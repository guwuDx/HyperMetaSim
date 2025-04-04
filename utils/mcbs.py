from DBUtils.PooledDB import PooledDB
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


    def check_duplicate_params(self, table:str, params:dict):
        query = f"SELECT ID FROM {table}\n"

        param_values = '\n  AND '.join(f"{k}={v}" for k, v in params.items())
        query += f"WHERE {param_values};"

        self.cursor.execute(query)
        res = self.cursor.fetchone()

        if res:
            logging.info(f"Duplicate parameters found in {table}: {res}")
            return res[0]
        else:
            return None


    def insert_params(self, table:str, params:dict, auto_commit=False):
        insert_sql = f"""
        INSERT INTO {table}"""

        columns = ', '.join(params.keys())
        values = ', '.join(['%s'] * len(params))

        insert_sql += f" ({columns}) VALUES ({values})"

        # logging.info(insert_sql)
        # logging.info(params.values())
        try:
            self.cursor.execute(insert_sql, tuple(params.values()))
            lastrowid = self.cursor.lastrowid
            if auto_commit:
                self.conn.commit()
            logging.info(f"Inserted parameters into {table} successfully.")
        except MySQLdb.Error as e:
            logging.error(f"Failed to insert parameters into {table}: {e}")
            if auto_commit:
                self.conn.rollback()
            return False
        return lastrowid


    def check_and_insert_param(self, table:str, params:dict):
        res = self.check_duplicate_params(table, params)
        if res:
            logging.info(f"Duplicate parameters found in {table}: {res}")
            return res
        else:
            logging.info(f"No duplicate parameters found in {table}")
            lastrowid = self.insert_params(table, params)
            logging.info(f"Parameters inserted into {table}, ID: {lastrowid}")
            return lastrowid


    def insert_freq_response(self, table:str, data:dict, foreign_key:int, auto_commit=False, force=False):
        duplicate_cnt = 0
        data_to_insert = []

        insert_sql = f"""
        INSERT INTO {table} (Frequency, real_pt, imag_pt, CP_ID)
        VALUES (%s, %s, %s, %s)
        """

        # data [[freq, real, imag], [freq, real, imag], ...]
        for freq, real, imag in data:
            query = f"""
            SELECT ID FROM {table}
            """ + """
              WHERE Frequency=%s
              AND real_pt=%s
              AND imag_pt=%s
              AND CP_ID=%s;
            """
            self.cursor.execute(query, (freq, real, imag, foreign_key))
            res = self.cursor.fetchone()
            if res:
                duplicate_cnt += 1
                # logging.info("x")
                if duplicate_cnt >= MAX_DUPLICATE_ALLOWED:
                    logging.warning(f"Too Many duplicate frequency response found in {table}: {res}")
                    if force:
                        duplicate_cnt = -1
                        logging.info(f"Force mode enabled, scanning for remaining data...")
                        continue
                    logging.info("Importing exiting data...")

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
                    continue
                continue
            else:
                duplicate_cnt = 0
                data_to_insert.append((freq, real, imag, foreign_key))
        # logging.info(insert)

        try:
            self.cursor.executemany(insert_sql, data_to_insert)
            if auto_commit:
                self.conn.commit()
            logging.info(f"Inserted all data ({len(data_to_insert)} records) successfully.")
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