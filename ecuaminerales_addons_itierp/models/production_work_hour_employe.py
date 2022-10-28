from email.policy import default

from odoo import api, fields, models
from odoo.addons.ecuaminerales_addons_itierp.models import production_work_hour
from odoo.exceptions import ValidationError


class ProductionWorkHourEmployee(models.Model):
    _name = 'production.work.hour.employee'
    _description = 'Modulo de Horas en producción'

    employee_id = fields.Many2one('hr.employee', string='Empleado')
    codigo_clock = fields.Integer("Código Relog", store=True, related='employee_id.codigo_clock', editable=True)
    resource_calendar_id = fields.Many2one('resource.calendar', 'Jornada de Trabajo', store=True,
                                           related='employee_id.resource_calendar_id')

    fecha_time = fields.Datetime("Marcación", store=True)
    hour = fields.Float("Hora", store=True)
    departamento = fields.Char("Departamento", store=True)
    dispositivo = fields.Char("Dispositivo", store=True)
    delete = fields.Boolean("Se Eliminara", store=True)
    festivo = fields.Boolean("Feriado?", store=True)
    dif = fields.Float("Diferencia en minutos", store=True, digits=(2, 6))
    dif_h = fields.Float("Diferencia en horas", store=True)
    production_work_hour = fields.Many2one("production.work.hour", "Horas de Producción", store=True,
                                           ondelete="cascade")
    type_mar = fields.Selection([('income', 'Ingreso'),
                                 ('exit', 'Salida'),
                                 ('old', 'Olvido'),
                                 ('error', 'Error')], default="error",
                                string="Tipo")
    turno = fields.Selection([('t1', 'Turno 1'),
                              ('t2', 'Turno 2'),
                              ('t3', 'Turno 3'),
                              ('t1f', 'Turno 1 F'),
                              ('t2f', 'Turno 2 F'),
                              ('t3f', 'Turno 3 F'),
                              ('tt2', 'Turno Doble'),
                              ('morning', 'Mañana'),
                              ('late', 'Tarde'),
                              ('seguido', 'Turno Unico'),
                              ('tt2', 'Turno Doble'),
                              ('no', 'SIN TURNO')], default="no",
                             string="Tipo")
