{
    'name': 'Certificate Verification',
    'version': '1.0',
    'summary': 'Certificate Verification',
    'sequence': 10,
    'description':  """This module will add a record to store report details""",
    'category':  """Tools""",
    'website': '',
    'depends': ['sale_management', ],
    'data': [
        'security/ir.model.access.csv',
        'views/certificate_views.xml',
        'views/sale_views.xml',
        'data/sequence.xml',
        'report/certificate.xml',

    ],
    'demo': [
    ],
    'installable': True,
    'application': True,
    'license': 'LGPL-3',
}