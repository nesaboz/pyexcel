{% extends 'pyexcel-setup.py.jj2' %}

{% block additional_keywords %}
    'tsv',
    'tsvz'
    'csv',
    'csvz',
    'xls',
    'xlsx',
    'ods'
{% endblock %}

{% block pyexcel_extra_classifiers %}
    'Development Status :: 3 - Alpha',
    'Programming Language :: Python :: Implementation :: PyPy'
{% endblock %}}
