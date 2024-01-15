import xlsxwriter
import pytz
import base64
from datetime import datetime,date,timedelta
from pytz import timezone
from io import BytesIO
from odoo import models,fields,api, _
import logging
logger = logging.getLogger(__name__)

class DateReportWizard(models.TransientModel):
    _name = 'kardex.general'
    _description = 'Date Report Wizard'

    date_from = fields.Date('Start Date', required=True)
    date_to = fields.Date('End Date', required=True)
    file_data = fields.Binary('File', readonly=True)
    product_id = fields.Many2one('product.template',string="Producto")

    @api.model
    def get_default_date_model(self):
        return pytz.UTC.localize(datetime.now()).astimezone(timezone('America/Lima'))
    
    def cell_format(self,workbook):
        cell_format={}
        cell_format['title']=workbook.add_format({
            'bold': True,
            'align': 'center',
            'valign': 'vcenter',
            'font_size': 25,
            'font_name': 'Arial',
        })
        cell_format['no'] = workbook.add_format({
            'align': 'center',
            'valign': 'vcenter',
            'border': False
        })
        cell_format['header']=workbook.add_format({
            'bold': True,
            'align': 'center',
            'valign': 'center',
            # 'bg_color':'#0BBB20',
            'border': True,
            'font_name': 'Arial'
        })
        cell_format['content'] = workbook.add_format({
            'font_size': 11,
            'border': False,
            'font_name': 'Arial'
        })
        cell_format['content_float'] = workbook.add_format({
            'font_size': 11,
            'border': False,
            'num_format': '#,##0.00',
            'font_name': 'Arial',
        })
        cell_format['total'] = workbook.add_format({
            'bold': True,
            # 'bg_color':'#CDCD08',
            'num_format': '#,##0.00',
            'align': 'center',
            'valign': 'center',
            'border': True,
            'font_name': 'Arial'
        })

        return cell_format,workbook
    

    def generate_excel_report(self):
        data=self.read()[0]
        date_from = data['date_from']
        date_to = data['date_to']
        product_id = data['product_id']
        # Crear un archivo Excel en memoria
        fp = BytesIO()
        workbook = xlsxwriter.Workbook(fp)
        cell_format, workbook = self.cell_format(workbook)
        report_name = _('Kardex Valorado')
        worksheet = workbook.add_worksheet(report_name)
        now = datetime.now() - timedelta(hours=5)

        columns = [
            _('FECHA'),
            _('DOCUMENTO'),
            _('PRODUCTO'),
            _('UBICACION'),
            _('CANTIDAD'),
            _('COSTO UNITARIO'),
            _('COSTO TOTAL'),
            _('CANTIDAD'),
            _('COSTO UNITARIO'),
            _('COSTO TOTAL'),
            _('CANTIDAD'),
            _('COSTO TOTAL'),
            _('COSTO UNITARIO')
        ]

        column_length = len(columns)
        if not column_length:
            return False
        
        no = 1
        column = 1

        worksheet.merge_range('A1:L1',_('REPORTE DE KARDEX VALORADO DESDE %s HASTA %s') %(date_from,date_to),cell_format['total'])
        worksheet.write('A4','No',cell_format['header'])
        # worksheet.merge_range('A3:A4','Fecha',cell_format['header'])
        # Definir formato para las celdas

        for col in columns:
            worksheet.write(3,column,col,cell_format['header'])
            column+=1
        
    # Stock_moves para el calculo del saldo Incial

        lines_si = None
        if product_id:
            # lines_si=self.env['stock.move'].search([('date','<',date_from),('product_id','=',product_id[0]),('state','=','done'),'|',('location_id.usage','=','internal'),('location_dest_id.usage','=','internal'),('location_dest_id.name','!=','location_id.name')])
            lines_si=self.env['stock.move'].search([('date','<',date_from),('product_id','=',product_id[0]),('state','=','done'),'|',('location_dest_id.usage','!=','internal'),('location_id.usage','!=','internal')])

        data_list_mov_si =[]
        data_list_si = {}
        data_list_si_ordenado = []

        for s in lines_si:
            costo_total=s.product_uom_qty*s.price_unit
            if s.location_id.usage =='internal':
                data_list_si = {
                    'date':s.date,
                    'reference':s.reference,
                    'product_id':s.product_id.name,
                    'location_dest_id':s.location_dest_id.name,
                    'product_uom_qty':s.product_uom_qty,
                    'price_unit':s.price_unit,
                    'costo_total':costo_total,
                }
            else:
                data_list_si = {
                    'date':s.date,
                    'reference':s.reference,
                    'product_id':s.product_id.name,
                    'location_id':s.location_id.name,
                    'product_uom_qty':s.product_uom_qty,
                    'price_unit':s.price_unit,
                    'costo_total':costo_total,
                }
            data_list_mov_si.append(data_list_si)
        
        data_list_si_ordenado = sorted(data_list_mov_si,key=lambda x: x['date'])

        logger.warning("------------DATOS DEL SALDO INICIAL---------------")
        logger.warning(data_list_mov_si)

        logger.warning("------------DATOS DEL SALDO INICIAL ORDENADO---------------")
        logger.warning(data_list_si_ordenado)

    # Calculo de los datos del Saldo Inicial

        cant_saldo_inicial = 0
        monto_saldo_inicial = 0
        for data_list_si in data_list_si_ordenado:
            precio_unit_si = data_list_si['price_unit']
            for clave, value in data_list_si.items():
                if clave == 'product_uom_qty':
                    cant_saldo_inicial = cant_saldo_inicial + value
                    monto_saldo_inicial = monto_saldo_inicial + precio_unit_si*value

        if cant_saldo_inicial > 0:
            costo_promedio = monto_saldo_inicial / cant_saldo_inicial
        else:
            costo_promedio = 0

        logger.warning("-----------SALDO INICIAL KARDEX--------------")
        logger.warning(cant_saldo_inicial)
        logger.warning(monto_saldo_inicial)
        logger.warning(costo_promedio)

    # Stock_moves para los movimientos

        data_list_mov = []
        data_list = {}
        data_list_mov_ordenado = []

        lines_kar = None
        if product_id:
            # lines_kar = self.env['stock.move'].search([('date','>=',date_from),('date','<=',date_to),('product_id','=',product_id[0]),('state','=','done'),'|',('location_id.usage','=','internal'),('location_dest_id.usage','=','internal'),('location_dest_id.name','!=','location_id.name')])
            # lines_kar = self.env['stock.move'].search([('date','>=',date_from),('date','<=',date_to),('product_id','=',product_id[0]),('state','=','done')])
            # lines_kar = self.env['stock.move'].search([('date','>=',date_from),('date','<=',date_to),('product_id','=',product_id[0]),('state','=','done'),('location_id.usage','=','internal'),('location_dest_id.usage','!=','internal'),'|',('location_id.usage','!=','internal'),('location_dest_id.usage','=','internal')])
            lines_kar = self.env['stock.move'].search([('date','>=',date_from),('date','<=',date_to),('product_id','=',product_id[0]),('state','=','done'),'|',('location_dest_id.usage','!=','internal'),('location_id.usage','!=','internal')])
            logger.warning("-----------MOVIMIENTOS OBTENIDOS--------------")
            logger.warning(lines_kar)


        for l in lines_kar:
            costo_total = l.product_uom_qty * l.price_unit
            if l.location_id.usage =='internal':
                data_list = {
                    'date':l.date,
                    'reference':l.reference,
                    'product_id':l.product_id.name,
                    'location_dest_id':l.location_dest_id.name,
                    'product_uom_qty':l.product_uom_qty,
                    'price_unit':l.price_unit,
                    'costo_total':costo_total,
                }
            else:
                data_list = {
                    'date':l.date,
                    'reference':l.reference,
                    'product_id':l.product_id.name,
                    'location_id':l.location_id.usage,
                    'product_uom_qty':l.product_uom_qty,
                    'price_unit':l.price_unit,
                    'costo_total':costo_total,
                }
            data_list_mov.append(data_list)

        data_list_mov_ordenado=sorted(data_list_mov,key=lambda x: x['date'])

        logger.warning("--------------LISTA DE MOVIMIENTOS-----------------")
        logger.warning(data_list_mov)

        logger.warning("--------------LISTA DE MOVIMIENTOS ORDENADOS-----------------")
        logger.warning(data_list_mov_ordenado)

        row = 5

        worksheet.write('L%s' % (row), cant_saldo_inicial,
                                cell_format['content_float'])
        worksheet.write('M%s' % (row), monto_saldo_inicial,
                                cell_format['content_float'])
        worksheet.write('N%s' % (row), costo_promedio,
                                cell_format['content_float'])

        column_float_number = {}

        cant_saldo = cant_saldo_inicial
        monto_saldo = monto_saldo_inicial

        for data_list in data_list_mov_ordenado:
            worksheet.write('A%s' % row, no, cell_format['no'])
            logger.warning("------------DATA IN DATA_LIST---------------")
            logger.warning(data_list)
            no += 1
            column = 1
            cont=0
            # for numero, (clave,valor) in enumerate(data_list.items(),start=1):
            for clave, value in data_list.items():

                if type(value) is int or type(value) is float:
                    content_format = 'content_float'
                    column_float_number[column] = column_float_number.get(
                        column, 0) + value
                    # logger.warning("-------------COLUMN_FLOAT_NUMBER----------------")
                    logger.warning(column_float_number)
                else:
                    content_format = 'content'

                if isinstance(value, datetime):
                    value = pytz.UTC.localize(value).astimezone(
                        timezone(self.env.user.tz or 'UTC'))
                    value = value.strftime('%Y-%m-%d %H:%M:%S')
                elif isinstance(value, date):
                    value = value.strftime('%Y-%m-%d')

                if cont==0:
                    worksheet.write(row, column, value,
                                cell_format[content_format])
                    if clave == 'product_uom_qty':
                        cant_saldo=cant_saldo+value
                        worksheet.write(row,column+6,cant_saldo,cell_format[content_format])
                    elif clave == 'costo_total':
                        monto_saldo = monto_saldo+value
                        worksheet.write(row,column+5,monto_saldo,cell_format[content_format])
                if column>=5 and cont>=1:
                    worksheet.write(row, column+3, value,
                                cell_format[content_format])
                    if clave == 'product_uom_qty':
                        cant_saldo=cant_saldo-value
                        worksheet.write(row,column+6,cant_saldo,cell_format[content_format])
                    elif clave == 'costo_total':
                        monto_saldo = monto_saldo-value
                        worksheet.write(row,column+5,monto_saldo,cell_format[content_format])
                    cont +=1

                if clave=='location_dest_id':
                    cont +=1
                column += 1

                logger.warning("-----------VALOR DE LISTA-------------")
                logger.warning(value)
                
            if cant_saldo>0:
                    costo_promedio=monto_saldo/cant_saldo
                    worksheet.write(row, column+5,costo_promedio)
            else:
                worksheet.write(row, column+5,0)
            
            row += 1

        # row -= 1

        for x in range(column_length + 1):

            if x == 0:
                worksheet.write('A%s' % (row + 1), _('Total'),
                                cell_format['total'])
            elif x not in column_float_number:
                worksheet.write(row, x, '', cell_format['total'])
            else:
                worksheet.write(
                    row, x, column_float_number[x], cell_format['total'])
        
        # Escribir encabezados

        # Obtener las fechas desde el asistente


        # Escribir las fechas en el informe


        # Ajustar el ancho de las columnas según el contenido

        # Cerrar el libro de trabajo y guardar en la memoria
        workbook.close()

        result = base64.encodebytes(fp.getvalue()).decode('utf-8')
        date_string = self.get_default_date_model().strftime("%Y-%m-%d")
        filename = '%s %s' % (report_name, date_string)
        filename += '%2Exlsx'
        self.write({'file_data': result})

        url = "web/content/?model=" + self._name + "&id=" + str(
            self[:1].id) + "&field=file_data&download=true&filename=" + filename

        # output.seek(0)
        return {
            'name': _('Generic Excel Report'),
            'type': 'ir.actions.act_url',
            'url': url,
            'target': 'new',
        }

