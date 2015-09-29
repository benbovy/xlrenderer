#!/usr/bin/env python
# -*- coding: utf-8 -*-


try:
    from setuptools import setup
except ImportError:
    from distutils.core import setup


with open('README.rst') as readme_file:
    readme = readme_file.read()

with open('HISTORY.rst') as history_file:
    history = history_file.read().replace('.. :changelog:', '')

requirements = [
    # TODO: put package requirements here
]

test_requirements = [
    # TODO: put package test requirements here
]

setup(
    name='xlrenderer',
    version='0.1.0',
    description="Render excel templates using a database and a specification file",
    long_description=readme + '\n\n' + history,
    author="Benoit Bovy",
    author_email='benbovy@gmail.com',
    url='https://github.com/benbovy/xlrenderer',
    packages=[
        'xlrenderer',
    ],
    package_dir={'xlrenderer':
                 'xlrenderer'},
    include_package_data=True,
    scripts=[
        'scripts/render_access2xls',
        'scripts/render_access2xls_gui'
    ],
    install_requires=requirements,
    license="ISCL",
    zip_safe=False,
    keywords='xlrenderer',
    classifiers=[
        'Development Status :: 2 - Pre-Alpha',
        'Intended Audience :: Developers',
        'License :: OSI Approved :: ISC License (ISCL)',
        'Natural Language :: English',
        "Programming Language :: Python :: 2",
        'Programming Language :: Python :: 2.6',
        'Programming Language :: Python :: 2.7',
        'Programming Language :: Python :: 3',
        'Programming Language :: Python :: 3.3',
        'Programming Language :: Python :: 3.4',
    ],
    test_suite='tests',
    tests_require=test_requirements
)
