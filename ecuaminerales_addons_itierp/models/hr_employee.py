# -*- coding: utf-8 -*-
# Part of Odoo. See LICENSE file for full copyright and licensing details.

from odoo import fields, models


class Employee(models.Model):
    _inherit = 'hr.employee'

    codigo_clock = fields.Integer("Código Reloj", store=True, help="Código unico asigando en relog")
    numero_marcaciones = fields.Integer("Numero de Marcaciones Diarias")
    _sql_constraints = [
        ('codigo_clock_uniq', 'unique (codigo_clock)', 'El código de reloj debe ser unico por empleado'),
    ]
