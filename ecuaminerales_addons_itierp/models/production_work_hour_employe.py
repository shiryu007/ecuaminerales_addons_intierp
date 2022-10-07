from email.policy import default

from odoo import api, fields, models
from odoo.addons.ecuaminerales_addons_itierp.models import production_work_hour
from odoo.exceptions import ValidationError


class ProductionWorkHourEmployee(models.Model):
    _name = 'production.work.hour.employee'
    _description = 'Modulo de Horas en producción'

    employee_id = fields.Many2one('hr.employee', string='Empleado')
    codigo_clock = fields.Integer("Código Relog", store=True, related='employee_id.codigo_clock', editable=True)
    fecha_time = fields.Datetime("Marcación", store=True)
    hour = fields.Float("Hora", store=True)
    departamento = fields.Char("Departamento", store=True)
    dispositivo = fields.Char("Dispositivo", store=True)
    delete = fields.Boolean("Se Eliminara", store=True)
    dif = fields.Float("Diferencia en minutos", store=True)
    production_work_hour = fields.Many2one("production.work.hour", "Horas de Producción", store=True,
                                           ondelete="cascade")
    type_mar = fields.Selection([('income', 'Ingreso'), ('exit', 'Salida'), ('error', 'Error')], default="error",
                                string="Tipo")
