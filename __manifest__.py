# -*- coding: utf-8 -*-
{
    'name': "Odoo Export Advanced",

    'summary': """
        Maach Odoo Dynamic Export""",

    'description': """
        This module provides a dynamic for records based on dynamic domains set
    """,

    'author': "Maduka Sopulu.",

    'category': 'Uncategorized',
    'version': '0.1',

    'depends': [
        'base',
    ],

    'data': [
        'security/access_groups.xml',
        'security/ir.model.access.csv',
        'views/odoo_export_views.xml',
    ],
    'demo': [
        # 'demo/demo.xml',
    ],
}
