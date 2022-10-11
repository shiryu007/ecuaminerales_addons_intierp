from datetime import datetime, timedelta, date

from odoo import api, fields, models
from odoo.exceptions import ValidationError
import base64
import xlrd
import calendar

# NUMERO MINUTOS DIFERENCIA
MINUTOS_DUPLICADO = 7
MINUTOS_TRABAJO = (8 * 60) + (2 * 60)


class ProductionWorkHour(models.Model):
    _name = 'production.work.hour'
    _description = 'Modulo de Horas en producción '
    _rec_name = 'sequence'

    sequence = fields.Char('Secuencia', required=False, readonly=True, track_visibility='onchange')
    document = fields.Binary('Documento de Horas', store=True)
    file_name = fields.Char('File Name', track_visibility='onchange')
    state = fields.Selection([('draft', 'Borrador'),
                              ('load', 'Cargado'),
                              ('purify', 'Depurado'),
                              ('confirm', 'Confirmar'),
                              ('posted', 'Pagado')],
                             readonly=True, default='draft', store=True, string="Estado", track_visibility='onchange')
    search_selection = fields.Selection([('code', 'Codigo Relog'), ('name', 'Nombre')],
                                        default='code', store=True, string="Ubicar empleado por",
                                        track_visibility='onchange')
    hour_production_ids = fields.One2many('production.work.hour.employee', 'production_work_hour', 'Lista de Horas')
    message = fields.Html("Mensaje de Error")
    register_count = fields.Integer('Numero de Registros', compute='_compute_count_registers')
    employee_search = fields.Many2one('hr.employee', 'Empleado')

    def _compute_count_registers(self):
        self.register_count = len(self.hour_production_ids)

    def view_registro_horas(self):
        self.ensure_one()
        return {
            'name': 'Registro de Horas',
            'type': 'ir.actions.act_window',
            'view_mode': 'tree,form',
            'res_model': 'production.work.hour.employee',
            'domain': [('id', 'in', self.hour_production_ids.ids)],
        }

    @api.model
    def create(self, vals):
        vals['sequence'] = self.env['ir.sequence'].next_by_code('production.work.hour.sequence') or '/'
        return super(ProductionWorkHour, self).create(vals)

    def load_information_of_file(self):
        if not self.document:
            raise ValidationError("Cargar Archivo de Horas")
        wb = xlrd.open_workbook(file_contents=base64.decodestring(self.document))
        sheet = wb.sheets()[0] if wb.sheets() else None
        data = [[sheet.cell_value(r, c) for c in range(sheet.ncols)] for r in range(sheet.nrows)]
        if data[0] == ['', '', '', '', '', '']:
            data.remove(data[0])
        if data[0] != ['Nombre', 'Número de empleado', 'Departamento', 'Fecha', 'Hora', 'Dispositivo']:
            raise ValidationError("""Recurde que el archivo debe contener la siguiente estructura  \n
            ['Nombre', 'Número de empleado', 'Departamento', 'Fecha', 'Hora', 'Dispositivo']""")
        self.hour_production_ids = False
        names_no_search = []
        name_not_range = []
        data_create = []
        for count, line in enumerate(data[1:]):
            name = line[0]
            employee = []
            if self.search_selection == 'name':
                employee = self.env['hr.employee'].search([('name', '=', line[0])])
            if self.search_selection == 'code':
                employee = self.env['hr.employee'].search([('codigo_clock', '=', int(line[1]))])
            if not employee:
                names_no_search.append(name)
                continue
            if employee and len(employee) > 1:
                employee = employee[0]
                name_not_range.append(name)
            data_create.append((0, count, {'employee_id': employee.id,
                                           'departamento': line[2],
                                           'fecha_time': self.conv_date_hout(line[3], line[4]),
                                           'hour': self.conv_time_float(line[4]),
                                           'dispositivo': line[5]}))
        self.hour_production_ids = data_create
        self.insert_messages(name_not_range, names_no_search)
        self.purge_data()
        self.state = 'load'

    def insert_messages(self, name_not_range, names_no_search):
        self.message = False
        if name_not_range or names_no_search:
            self.message = "<ul>"
        else:
            self.message = False
        for name in set(name_not_range):
            self.message += "<li class='text-danger'> Empleado: %s multiples coincidencias </li> \n " % name
        for name in set(names_no_search):
            self.message += "<li class='text-warning'>  Empleado: %s no encontrado </li> \n " % name

        if self.message:
            self.message += "</ul>"

    def conv_time_float(self, value):
        vals = value.split(':')
        t, hours = divmod(float(vals[0]), 24)
        t, minutes = divmod(float(vals[1]), 60)
        minutes = minutes / 60.0
        return float(hours + minutes)

    def conv_date_hout(self, date, time):
        date_time_str = date + " " + time + ':00'
        fecha = datetime.strptime(str(date_time_str), '%m/%d/%Y %H:%M:%S')
        fecha = fecha + timedelta(hours=5)
        return fecha

    def purge_data(self):
        if not self.hour_production_ids:
            return True
        self.hour_production_ids.write({'delete': False, 'type_mar': 'error', 'dif': 0, 'turno': 'no'})
        for employee_id in set(self.hour_production_ids.mapped('employee_id')):
            list_hours = self.hour_production_ids.filtered(lambda x: x.employee_id == employee_id).sorted('fecha_time')
            count = 1
            for ahora in list_hours[1:]:
                antes = list_hours[count - 1]
                diferencia = ahora.fecha_time - antes.fecha_time
                minutes = abs(diferencia.total_seconds() / 60)
                ahora.dif = minutes
                ahora.dif_h = minutes / 60
                if minutes < MINUTOS_DUPLICADO:
                    ahora.delete = True
                    ahora.type_mar = antes.type_mar
                    ahora.turno = antes.turno
                    if ahora.type_mar == 'exit':
                        antes.delete = True
                        ahora.delete = False
                else:
                    self.detectar_ingreso_salida(antes, ahora, minutes)
                count += 1

    def detectar_ingreso_salida(self, antes, ahora, minutes):
        # TURNO ROTATIVOS
        sales_journal_id = self.env.ref('ecuaminerales_addons_itierp.resource_rotativos')
        if ahora.resource_calendar_id == sales_journal_id:
            # VALIDAR HORARIOS
            # 	             |      TURNO 1	   | TURNO 2	        | TURNO 3
            # LUNES-VIERNES	 | 06H00 - 14H00   |	14H00 - 22H00	| 22H00 - 06H00
            # SÁBADO	     |06H00 - 18H00	   |    LIBRE	        | 18H00 - 06H00
            # DOMINGO	     |06H00 - 18H00    |	18H00 - 06H00	| LIBRE
            f_antes = antes.fecha_time - timedelta(hours=5)
            f_ahora = ahora.fecha_time - timedelta(hours=5)
            if f_antes.day == 26:
                print("To ca ver que pasa aca")
            if antes.turno != 'no':
                return True
            if f_antes.weekday() in [calendar.MONDAY, calendar.TUESDAY, calendar.WEDNESDAY, calendar.THURSDAY,
                                     calendar.FRIDAY]:
                if (minutes / 60) > 14:
                    antes.type_mar = 'old'
                    return True
                if 5 <= f_antes.hour <= 7 and f_ahora.hour <= 18:
                    antes.type_mar = 'income'
                    ahora.type_mar = 'exit'
                    antes.turno = 't1'
                    ahora.turno = 't1'
                    return True
                if 13 <= f_antes.hour <= 15 and f_ahora.hour <= 24 or 13 <= f_antes.hour <= 15 and f_ahora.hour <= 2:
                    antes.type_mar = 'income'
                    ahora.type_mar = 'exit'
                    antes.turno = 't2'
                    ahora.turno = 't2'
                    return True
                if 21 <= f_antes.hour <= 23 and f_ahora.hour <= 8:
                    antes.type_mar = 'income'
                    ahora.type_mar = 'exit'
                    antes.turno = 't3'
                    ahora.turno = 't3'
                    return True
                if 9 <= f_antes.hour <= 11 and f_ahora.hour <= 23:
                    antes.type_mar = 'income'
                    ahora.type_mar = 'exit'
                    antes.turno = 'tt2'
                    ahora.turno = 'tt2'
                    return True

            if f_antes.weekday() in [calendar.SATURDAY, calendar.SUNDAY]:
                if 5 <= f_antes.hour <= 7 and f_ahora.hour <= 20:
                    antes.type_mar = 'income'
                    ahora.type_mar = 'exit'
                    antes.turno = 't1f'
                    ahora.turno = 't1f'
                    return True

            if f_antes.weekday() in [calendar.SATURDAY]:
                if 15 <= f_antes.hour <= 19 and f_ahora.hour <= 8:
                    antes.type_mar = 'income'
                    ahora.type_mar = 'exit'
                    antes.turno = 't2f'
                    ahora.turno = 't2f'
                    return True
            if f_antes.weekday() in [calendar.SUNDAY]:
                if 15 <= f_antes.hour <= 19 and f_ahora.hour <= 8:
                    antes.type_mar = 'income'
                    ahora.type_mar = 'exit'
                    antes.turno = 't3f'
                    ahora.turno = 't3f'
                    return True
            print("Toca ver por que llegas hasta aca....")

    def delete_duplicates(self):
        self.hour_production_ids = self.hour_production_ids.filtered(lambda x: not x.delete)
        self.purge_data()
