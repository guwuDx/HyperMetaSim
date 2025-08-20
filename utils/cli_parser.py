import argparse


def _list_parser(option_params):
    """
    List parser for command line options.
    :param option_params: List of tuples containing option name and description.
    :return: List of argparse arguments.
    """
    items = []
    return [x.strip() for x in option_params.split()]

def parser_constructor():
    """
    Constructs the argument parser for the command line interface.
    :return: Argument parser object.
    """
    psr = argparse.ArgumentParser(prog="hms" , description="HyperMetaSim - A CST-based Automatic simulation tool for metamaterials.")
    sub_psrs = psr.add_subparsers(dest="mode", description="work mode")

    # sub parser: http
    psr_http = sub_psrs.add_parser("http", help="http simulate (web interface)")


    # sub parser: playbook
    psr_pb = sub_psrs.add_parser("pb", help="playbook scripts, to run a predefined sequence of operations")


    # sub parser: cli
    psr_cli = sub_psrs.add_parser("cli", help="command line interface")


    # sub parser: sql
    psr_sql = sub_psrs.add_parser("sql", help="database options")
    sub_sql_psrs = psr_sql.add_subparsers(dest="action", help="db tasks")

    # # subsub parser: check duplicate
    psr_sql_chk = sub_sql_psrs.add_parser("check-dup", help="check duplicate data")
    psr_sql_chk.add_argument("--check-only", action="store_true",
                             help="check only, do not delete")
    psr_sql_chk.add_argument("-o", "--output", type=str, default="./duplicate_data.txt",
                             help="output report for duplicate data")

    # # subsub parser: import data
    psr_sql_import = sub_sql_psrs.add_parser("cst2mysql", help="simulate result import")
    psr_sql_import.add_argument("-p", "--path", type=str, required=True,
                                 help="Path to data file/directory for project being processed.")
    psr_sql_import.add_argument("-s", "--sparam-names", nargs="+", type=_list_parser, 
                                 help="specific list of s-parameter names to import (separated by space), null for all")
    psr_sql_import.add_argument("-c", "--compensation", type=float, default=0,
                                 help="compensation the length of substrate")
    psr_sql_import.add_argument("--file-parallel", type=int, default=1,
                                 help="number of file parallelism, " \
                                 "  0:    auto" \
                                 "  1:    sequential" \
                                 "  > 1:  parallel")
    psr_sql_import.add_argument("--thread", type=int, default=0,
                                 help="number of threads (max worker)" \
                                 "  0|1:  sequential" \
                                 "  > 1:  parallel")
    psr_sql_import.add_argument("-f", "--force", action="store_true", 
                                 help="force import")
    psr_sql_import.add_argument("--check-after-import", action="store_true",
                                 help="check data after import")


    # sub parser: debug
    psr_deb = sub_psrs.add_parser("debug", help="execute debug content")


    return psr.parse_args()