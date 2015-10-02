# -*- coding: utf-8 -*-

"""
Utils: database connections/engines and jinja2 custom filters.

"""

import numpy as np
import sqlalchemy
import pypyodbc
import jinja2


__all__ = ["connect_access_db", "create_access_engine", "jinja_custom_env"]


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


def none2empty_filter(val):
    """Jinja2 template to convert None value to empty string."""
    if not val is None:
        return val
    else:
        return ''

def nan2empty_filter(val):
    """Jinja2 template to convert 'nan' value to empty string."""
    try:
        if np.isnan(val):
            return ''
    except TypeError:
        pass
    return val

jinja_custom_env = jinja2.Environment()
jinja_custom_env.filters['none2empty'] = none2empty_filter
jinja_custom_env.filters['nan2empty'] = nan2empty_filter
