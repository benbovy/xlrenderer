# -*- coding: utf-8 -*-

import sqlalchemy
import pyodbc
import pypyodbc


def connect_access_db(filename):
    """Returns a connection to the Access database."""
    
    driver = '{Microsoft Access Driver (*.mdb, *.accdb)}'
    con = pypyodbc.connect('Driver={0};Dbq={1};Uid=Admin;Pwd=;'
                           .format(driver, filename))
    return con

def create_access_engine(filename):
    """
    Creates a new SQLAlchemy engine from an Access
    database (.mdb, .accdb).
    """
    engine = sqlalchemy.create_engine(
        'mysql+pyodbc://',
        creator=lambda: connect_access_db(filename)
    )
    return engine
