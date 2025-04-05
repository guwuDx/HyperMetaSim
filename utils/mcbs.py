from DBUtils.PooledDB import PooledDB
import uuid6
import json
import MySQLdb
import logging
import numpy as np

from utils.misc import read_config

MAX_DUPLICATE_ALLOWED = 10


class SQLPool:
    _instance = None

    def __new__(cls):
        mysql_config = read_config("./config/service.json", "mysql")
        if not cls._instance:
            cls._instance = super(SQLPool, cls).__new__(cls)
            cls._instance._pool = PooledDB(
                creator = MySQLdb,
                maxconnections = mysql_config["pool"]["max_connections"],
                mincached = mysql_config["pool"]["mincached"],
                maxcached = mysql_config["pool"]["maxcached"],
                blocking = mysql_config["pool"]["blocking"],
                host = mysql_config["host"],
                port = mysql_config["port"],
                user = mysql_config["user"],
                passwd = mysql_config["password"],
                db = mysql_config["database"],
                charset = "utf8mb4",
            )
        return cls._instance
    
    def get_connection(self):
        return self._pool.connection()


class SQLHandler:
    def __init__(self):
        self.pool = SQLPool()
        self.conn = self.pool.get_connection()
        self.cursor = self.conn.cursor()


    def check_duplicate_params(self, table: str, params: dict):
        """
        Check if a row with the same parameter values exists in the given table.

        :param table: Table name. Example: 'CuboidPillar_parameters'.
        :param params: dict of {column_name: value}, where value can be int, float, str, bytes(UUID), etc.
        :return: The 'ID' if a duplicate row is found, otherwise None.
        """

        # Build WHERE clause with placeholders (e.g. "col1=%s AND col2=%s ...")
        where_clause = ' AND '.join(f"{k}=%s" for k in params.keys())
        query = f"SELECT ID FROM {table} WHERE {where_clause};"

        try:
            # Use tuple(...) to pass the values so that it won't raise 'not all arguments converted' error
            self.cursor.execute(query, tuple(params.values()))
            res = self.cursor.fetchone()
            if res:
                # If a duplicate is found, return the ID
                return res[0]
            else:
                # No duplicate found
                return None
        except MySQLdb.Error as e:
            logging.error(f"[ERROR] Failed to check duplicates in {table}: {e}")
            return None


    def insert_params(self, table:str, params:dict, 
                      uuid=True, auto_commit=False):
        """
        Insert a new record into the specified table with the given parameters.

        :param table: The table name. For example 'generic_parameters' or 'CuboidPillar_parameters'.
        :param params: dict of {column_name: value}, which can include bytes (for UUID) or standard numeric types.
        :param uuid: Whether to auto-generate a UUID for the 'ID' column if not provided.
        :param auto_commit: Whether to auto commit transaction after insertion.
        :return: The lastrowid if insertion succeeds, or False on error.
        """
        if 'ID' not in params and uuid:
            # Generate a new UUID for ID if not provided
            new_uuid6 = uuid6.uuid6()
            params['ID'] = new_uuid6.bytes
            logging.debug(f"Auto-generated UUIDv6 for table={table}: {new_uuid6} (bytes={new_uuid6.bytes})")

        # Build INSERT statement with placeholders
        columns = ', '.join(params.keys())
        placeholders = ', '.join(['%s'] * len(params))
        insert_sql = f"INSERT INTO {table} ({columns}) VALUES ({placeholders})"

        # logging.info(insert_sql)
        # logging.info(params.values())
        try:
            # Execute with parameter binding
            self.cursor.execute(insert_sql, tuple(params.values()))
            lastrowid = self.cursor.lastrowid

            if auto_commit:
                self.conn.commit()
            logging.info(f"Inserted parameters into {table} successfully. lastrowid={lastrowid}")

            # If it's an auto-increment table, lastrowid is the real ID.
            return params.get('ID', lastrowid)

        except MySQLdb.Error as e:
            logging.error(f"Failed to insert parameters into {table}: {e}")
            if auto_commit:
                self.conn.rollback()
            return False


    def check_and_insert_param(self, table:str, params:dict, return_duplicate_status=False):
        """
        Check duplicates and insert if none.

        :param table: The table name.
        :param params: The parameter dict for insertion.
        :return: ID of the existing or newly inserted row, or False if insertion fails.
        """
        res = self.check_duplicate_params(table, params)
        if res:
            logging.info(f"Duplicate parameters found in {table}: {res}")
            if return_duplicate_status:
                return res, True
            else:
                return res
        else:
            logging.info(f"No duplicate parameters found in {table}, inserting new row.")
            lastrowid = self.insert_params(table, params)
            if lastrowid is not False:
                logging.info(f"Parameters inserted into {table}, ID: {lastrowid}")
            if return_duplicate_status:
                return lastrowid, False
            else:
                return lastrowid


    def insert_freq_response(self, table:str, data:dict, gp_id:int, shp_id:int,
                             auto_commit=False, force=False):
        """
        Insert frequency response data into the specified freq_response table, with duplication checks.

        :param table: e.g. 'CuboidPillar_freq_resp_NIR'
        :param data: list of [freq, real_pt, imag_pt], e.g. [[33.1, 0.25, 0.39], [33.2, 0.26, 0.40], ...]
        :param foreign_key: The CP_ID or relevant BINARY(16) ID (or int ID if table expects that).
        :param auto_commit: Whether to commit automatically.
        :param force: If True, continue inserting despite duplicates (except hitting MAX_DUPLICATE_ALLOWED).
        :return: True if all inserted, False if error or too many duplicates found.
        """
        duplicate_cnt = 0
        data_to_insert = []

        insert_sql = f"""
        INSERT INTO {table} (ID, Frequency, real_pt, imag_pt, GP_ID, ShP_ID)
        VALUES (%s, %s, %s, %s, %s, %s)
        """

        # data [[freq, real, imag], [freq, real, imag], ...]
        for freq, real_val, imag_val in data:
            query = f"""
            SELECT ID FROM {table}
            """ + """
              WHERE Frequency=%s
                AND GP_ID=%s
                AND ShP_ID=%s;
            """
            try:
                self.cursor.execute(query, (freq, gp_id, shp_id))
                res = self.cursor.fetchone()
            except MySQLdb.Error as e:
                logging.error(f"[ERROR] Querying duplicates in {table} failed: {e}")
                if auto_commit:
                    self.conn.rollback()
                return False

            if res:
                duplicate_cnt += 1
                # logging.info("x")
                if duplicate_cnt >= MAX_DUPLICATE_ALLOWED:
                    logging.warning(f"Too Many duplicate frequency response found in {table}: {res}")
                    if force:
                        # Force mode: continue scanning rest
                        duplicate_cnt = -1
                        logging.info("Force mode enabled, ignoring duplicates for the rest data.")
                        continue

                    logging.info("No force mode => partial import of existing data_to_insert.")
                    if data_to_insert:
                        try:
                            self.cursor.executemany(insert_sql, data_to_insert)
                            if auto_commit:
                                self.conn.commit()
                            logging.info(f"Inserted partial data ({len(data_to_insert)} records).")
                        except MySQLdb.Error as e:
                            logging.error(f"Failed to insert data into {table}: {e}")
                            if auto_commit:
                                self.conn.rollback()
                            return False
                        return False
                    else:
                        logging.info(f"Duplicate data found, but no data to insert.")
                        if auto_commit:
                            self.conn.rollback()
                        return False
                elif duplicate_cnt < 0:
                    # Means we are in 'force' ignoring duplicates
                    continue
                else:
                    # Normal duplicate, just skip
                    continue
            else:
                duplicate_cnt = 0
                uuid = uuid6.uuid6().bytes
                data_to_insert.append((uuid, freq, real_val, imag_val, gp_id, shp_id))

        # After loop, insert the leftover data
        try:
            if data_to_insert:
                self.cursor.executemany(insert_sql, data_to_insert)
                if auto_commit:
                    self.conn.commit()
                logging.info(f"Inserted all data ({len(data_to_insert)} records) successfully in {table}.")
            else:
                logging.info("No data to insert or all duplicates found.")
            return True
        except MySQLdb.Error as e:
            logging.error(f"Failed to insert data into {table}: {e}")
            if auto_commit:
                self.conn.rollback()
            return False


    def load_shape_definition(self, shape_name:str):
        select_sql = """
        SELECT * FROM shape_def
        """

        if shape_name:
            select_sql += f"WHERE name='{shape_name}'"
        # select_sql += ";"
        self.cursor.execute(select_sql)
        row = self.cursor.fetchone()

        if not row:
            logging.error(f"Shape {shape_name} not found in shape_def table")
            return None
        
        # construct the dictionary structure
        shape_info = {
            "ID": row[0],
            "name": row[1],
            "zn": row[2],
            "parameters_num": row[3],
            "col_names": json.loads(row[4]),
            "tables": json.loads(row[5])
        }
        return shape_info


    def load_generic_parameters_definition(self):
        return self.load_shape_definition("generic_parameters")


    def get_param_info(self, shape_name:str):
        generic_info = self.load_generic_parameters_definition()
        shape_info = self.load_shape_definition(shape_name)

        filtered_param_info = []
        param_names = generic_info["col_names"] + shape_info["col_names"]
        param_info = read_config("config/ParamInfo.json")
        for info in param_info:
            if info["name"] in param_names and info["belongs_to"] in (shape_name, "generic_parameters"):
                filtered_param_info.append(info)

        return filtered_param_info


    def get_sparam_id(self, sparam_name:str):
        # delete first char of sparam_name
        if sparam_name[0] == "S":
            sparam_name = sparam_name[1:]

        query = """
        SELECT ID FROM s_parameter
        WHERE S_Param=%s
        """
        self.cursor.execute(query, (sparam_name,))
        res = self.cursor.fetchone()
        if res:
            return res[0]
        else:
            logging.error(f"S-parameter {sparam_name} not found in sparam_def table")
            return None


    def get_material_id(self, material_identifier:str):
        query = """
        SELECT ID FROM material_def
        WHERE identifier=%s
        """

        self.cursor.execute(query, (material_identifier,))
        res = self.cursor.fetchone()
        if res:
            return res[0]
        else:
            logging.error(f"Material {material_identifier} not found in material_def table")
            return None


    def commit(self):
        self.conn.commit()


    def rollback(self):
        self.conn.rollback()


    def close(self):
        self.cursor.close()
        self.conn.close()