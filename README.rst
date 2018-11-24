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

.. code-block:: python

    >>> from xlrenderer import ExcelTemplateRenderer
    >>> from sqlalchemy import create_engine
    >>> engine = create_engine('protocol://user@localhost:port/mydatabase')
    >>> xltemplate = "/path/to/excel/template.xlsx"
    >>> def_file = "/path/to/database2excel.yml"
    >>> outdir = "/path/to/outputdir"
    >>> r = ExcelTemplateRenderer(engine, xltemplate, def_file, outdir)
    >>> r.render()

Taking the contacts table below as an example (stored in a relational database),

+----+------------+-----------+------------+--------+
| id | first_name | last_name | birth_date | gender |
+====+============+===========+============+========+
| 1  | Anna       | Harper    | 04-03-1982 | F      |
+----+------------+-----------+------------+--------+
| 2  | Fred       | Lloyd     | 10-12-1976 | M      |
+----+------------+-----------+------------+--------+
| 3  | John       | Doe       | 22-06-1965 | M      |
+----+------------+-----------+------------+--------+
| 4  | Daisy      | Schaefer  | 08-09-1989 | F      |
+----+------------+-----------+------------+--------+

A simple YAML definition block would look like

.. code-block:: yaml

    - name: simple contact table
      query: >
        SELECT * FROM [CONTACTS]
      apply_by_row: no
      cell_specification:
        worksheet: "Contacts"
        top_left_cell: A1
        header: yes
        index: no
      save_as:
        filename: "contacts.xlsx"
        export_pdf: yes

where ``name`` is any name given to the definition block (see below), ``query`` is the SQL query used to get data from the database and render it in the Excel template, ``apply_by_row: no`` here means that the whole query result will be rendered as a table of contiguous cells in the xls file, and the ``cell_specification`` block is where we define the name of the worksheet, the top-left cell of the rendered table and whether or not to show the header (i.e., field names) and the index (here the ``id`` key). Finally, the ``save_as`` block allows to save the rendered template in a separate file, with an option to also export it as PDF. 

More advanced rendering is possible. For example, the template might here consist of a custom contact form (non-contiguous cells) to be filled and rendered for each person in separate files. The corresponding YAML definition block would then look like

.. code-block:: yaml

    - name: custom contact form
      query: >
        SELECT * FROM [CONTACTS]
      apply_by_row: yes
      cell_specification:
        worksheet: "Contact Info"
        cells:
          - { cell: B2, content: "{{ first_name|capitalize }}" }
          - { cell: B3, content: "{{ last_name|capitalize }}" }
          - { cell: C6, content: "{{ birth_date.strftime('%d/%m/%Y') }}" }
          - { cell: E6, content: "{% if gender == 'M' %}X{% endif %}" }
          - { cell: E7, content: "{% if gender == 'F' %}X{% endif %}" }
      save_as:
        filename: "{{ first_name }}-{{ last_name }}.xlsx"
        export_pdf: yes

Note ``apply_by_row: yes`` which will fill, render and export the template for each row of the query result. Note also the use of jinja2's templating language for the cell content and filename.

For even more advanced rendering, it is possible to combine multiple definition blocks (with results from different queries) using ``include``, e.g.,

.. code-block:: yaml

    - name: custom contact form
      query: >
        SELECT * FROM [CONTACTS]
      include:
        - name of another definition block
      save_as:
        filename: "{{ first_name }}-{{ last_name }}.xlsx"
        export_pdf: yes

License
-------

Copyright (c) 2015-2018 Benoit Bovy.

Licensed under the terms of the MIT License
