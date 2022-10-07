# Copyright 2021 La Colina.
# License LGPL-3.0 or later (http://www.gnu.org/licenses/lgpl.html).

{
    "name": "Ecuaminerales Addons Itierp",
    "summary": "production_work_hour",
    'description': """ Modulo para calculo de horas en producción """,
    "version": "0.0.1",
    "category": "Industries",
    "website": 'https://intierp.com/',
    "author": "Ricardo Vinicio Jara Jara | "
              "Itierp",
    "license": "LGPL-3",
    "installable": True,
    "depends": [
        "base", 'hr_attendance'
    ],
    "data": [
        'security/ir.model.access.csv',
        'data/data.xml',
        'data/resource_data.xml',
        'views/production_work_hour_menu.xml',
        'views/hr_employee_view.xml',
        'views/view_production_work_hour.xml',
        'views/view_production_work_employee.xml'
    ],
    'js': ['static/src/js/ecuaminerales_addons_itierp.js'],
    'css': ['static/src/css/ecuaminerales_addons_itierp.scss']
}
