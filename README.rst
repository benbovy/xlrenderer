===============================
xlrenderer
===============================

Populate and render Excel templates from any database, using a single
YAML definition file.


Features
--------

Given any database supported by `sqlalchemy` and a "template" Excel file, this
package allows to generate (many) data-populated Excel files (and PDFs,
Windows only), according to user-defined queries and worksheet cell 
location/content specified in a definition file.

The structure and format (YAML) of the definition file allow to render
complex templates with little effort. The jinja2 templating language is used
for easy content rendering.


Installation
------------

This package is only available for Windows and OSX platforms with
Excel installed.

Given that all requirements below are satisfied, run:

    $ python setup.py install

No package available yet on PyPI or Anaconda.org.


Requirements
------------

- pandas
- sqlalchemy
- xlwings
- pyyaml
- jinja2


Usage
-----

Basic usage from within Python:

    >>> from xlrenderer import ExcelTemplateRenderer
    >>> from sqlalchemy import create_engine
    >>> engine = create_engine('protocol://user@localhost:port/mydatabase')
    >>> xltemplate = "/path/to/excel/template.xlsx"
    >>> def_file = "/path/to/database2excel.yml"
    >>> outdir = "/path/to/outputdir"
    >>> r = ExcelTemplateRenderer(engine, xltemplate, def_file, outdir)
    >>> r.render()

TODO: document the definition file (provide an example).


License
-------

Copyright (c) 2015 Benoit Bovy.

Licensed under the terms of the MIT License
