# -*- coding: utf-8 -*-

import os
import logging

import pandas as pd

from xlwings import Workbook, Sheet, Range

import yaml
import jinja2


logger = logging.getLogger('xlrenderer')


class ExcelTemplateRenderer(object):
    """
    A class to render an Excel template using data stored
    in a database and given a specification file.
    
    Parameters
    ----------
    db_engine : object
        the :class:`sqlalchemy.engine.Engine` object used to
        connect and query the database.
    template_name : str
        path to the `xlsx` template file.
    spec_filename : str
        path to the specification file (yaml format).
    output_dirname : str
        path to the directory where to store the
        generated excel file(s) (i.e., the rendered template(s)).
    jinja_env : object or None
        allow to provide a :class:`jinja2.Environment` object,
        helpful when using custom filters in the specification file
        (optional).
    
    Notes
    -----
    The specification file (yaml format) consists of a list
    of render blocks. TODO: documentation on the format.

    """
    
    def __init__(self, db_engine, template_name,
                 spec_filename, output_dirname, jinja_env=None):
        
        self.db_engine = db_engine
        self.template_name = template_name
        self.spec_filename = spec_filename
        self.output_dirname = os.path.abspath(output_dirname)
        
        with open(self.spec_filename, 'r') as f:
            self.render_blocks = yaml.load(f)
        
        os.makedirs(self.output_dirname, exist_ok=True)
        logger.info("output directory is %s", self.output_dirname)
    
    def open_template_as_current_wkb(self):
        self.wkb = Workbook(
            os.path.abspath(self.template_name),
            app_visible=False
        )
        self.wkb.set_current()
    
    def save_and_close_current_wkb(self, filename):
        filepath = os.path.join(self.output_dirname, filename)
        self.wkb.save(filepath)
        
        logger.info("created %s", filepath)

        self.wkb.close()

    def insert_one_series(self, series, cell_specification):
        """
        Populate the current workbook given a single
        :class=:`pandas.Series` object.
        """
        if not len(series):
            return
        
        # contiguous cells
        #TODO: (use vertical and horizontal properties of xlwings)
        
        # non-contiguous user-defined cells
        for cs in cell_specification.get('cells', []):
            ws = cs.get('worksheet') or Sheet.active(self.wkb).name
            content = jinja2.Template(cs['content']).render(**series)
            
            logger.debug("insert content '%s' at cell '%s' in sheet '%s'",
                         content, cs['cell'], ws)

            Range(ws, cs['cell']).value = content
    
    def insert_one_dataframe(self, df, cell_specification):
        """
        Populate the current workbook given a single
        :class=:`pandas.DataFrame` object.
        """
        if not len(df):
            return

        index = cell_specification.get('index', False)
        header = cell_specification.get('header', False)
        top_left_cell = cell_specification.get('top_left_cell', 'A0')
        
        logger.debug("insert %d by %d rows/cols dataframe "
                     "at cell '%s' in sheet '%s'",
                     len(df), len(df.columns),
                     str(top_left_cell), Sheet.active(self.wkb).name)
        
        Range(top_left_cell, index=index, header=header).value = df

    def apply_render_block(self, render_block, query_context=None,
                           **kwargs):
        """
        Apply a single render block in the specification file.

        - `query_context` (mappable or None) is a context used when
          rendering the database query with jinja2 (optional). 
        - **kwargs is used to overwrite any key/value pair
          in the render block.

        """
        # override render_block key/val with kwargs
        render_block.update(kwargs)

        # query the DB into a pandas DataFrame
        if query_context is None:
            query_context = dict()
        query_template = jinja2.Template(render_block['query'].strip())
        query = query_template.render(**query_context)
        logger.debug("rendered query: \n'''\n%s\n'''", query)
        df = pd.read_sql(query, self.db_engine)
        
        logger.debug("query returned %d record(s)", len(df))

        # TODO: calculate extra columns and add it to the DataFrame

        # activate worksheet if provided
        ws_name = render_block['cell_specification'].get('worksheet') or None
        if ws_name is not None:
            ws2reactivate_name = Sheet.active(self.wkb).name
            Sheet(ws_name, wkb=self.wkb).activate()

        # apply the render_block, apply recusively included blocks,
        # and save the rendered workbook(s) if needed
        apply_by_row = render_block.get('apply_by_row', False)
        save_as = render_block.get('save_as', None)
        
        if apply_by_row and save_as is not None:
            logger.info("%d file(s) to generate", len(df))
        
        if apply_by_row:
            for row, pseries in df.iterrows():
                print(Sheet.active(self.wkb).name)
                self.insert_one_series(
                    pseries, render_block['cell_specification']
                )
            
                for item in render_block.get('include', []):
                    if isinstance(item, dict):
                        block_name = item.pop('render_block')
                        override_vars = item
                    else:
                        block_name = item
                        override_vars = {}
                    block = [b for b in self.render_blocks
                             if b['name'] == block_name][0]
                    self.apply_render_block(block,
                                            query_context=pseries,
                                            **override_vars)
                
                if save_as is not None:
                    filename = jinja2.Template(save_as).render(**pseries)
                    self.save_and_close_current_wkb(filename)
                    self.open_template_as_current_wkb()
                    # re-activate the sheet because re-opened the template                    
                    if ws_name is not None:
                        Sheet(ws_name, wkb=self.wkb).activate()
        
        else:
            self.insert_one_dataframe(
                df, render_block['cell_specification']
            )
            # TODO: include and save_as in this case
        
        # re-activate former worksheet if needed
        if ws_name is not None:
            Sheet(ws2reactivate_name, wkb=self.wkb).activate()

    def render(self):
        """Main render method."""
        
        self.open_template_as_current_wkb()
        
        save_render_blocks = [block for block in self.render_blocks
                              if 'save_as' in block.keys()]
        
        for block in save_render_blocks:
            self.apply_render_block(block)
        
        self.wkb.close()
