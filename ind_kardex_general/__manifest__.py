{
    'name':'Indomin Kardex General',
    'description':'Movimientos basados en las entradas y salidas no considera las transferencias',
    'author':'Juan Carlos',
    'depends':['stock','base','report_xlsx'],
    'data':[
            'security/ir_model_access.xml',
            'models/kardex_general.xml',
            ]
}