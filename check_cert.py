import OpenSSL
from asn1crypto import x509
import base64


def get_cert_attr(cert_path):
    result = {}
    cert_obj = open(cert_path, 'rb').read()
    try:
        cert = OpenSSL.crypto.load_certificate(OpenSSL.crypto.FILETYPE_ASN1, cert_obj)
    except OpenSSL.crypto.Error:
        cert = OpenSSL.crypto.load_certificate(OpenSSL.crypto.FILETYPE_PEM, cert_obj)

    not_after = cert.get_notAfter().decode('ascii') # 20211029094827
    result['конец'] = f'{not_after[6:8]}.{not_after[4:6]}.{not_after[:4]} {not_after[8:10]}:{not_after[10:12]}:{not_after[12:14]}'
    not_before = cert.get_notBefore().decode('ascii')
    result['начало'] = f'{not_before[6:8]}.{not_before[4:6]}.{not_before[:4]} {not_before[8:10]}:{not_before[10:12]}:{not_before[12:14]}'
    result['Серийный номер'] = '{0:x}'.format(cert.get_serial_number())
    subject_ext = cert.get_subject().get_components()

    sn = ''
    gn = ''
    l = ''
    street = ''
    for key, value in cert.get_subject().get_components():
        #print(key.decode('utf-8'), value.decode('utf-8'))
        if key.decode('utf-8') == 'CN':
            result['Общее имя'] = value.decode('utf-8')
        if key.decode('utf-8') == 'SN':
            sn = value.decode('utf-8')
        if key.decode('utf-8') == 'GN':
            gn = value.decode('utf-8')
        if sn and gn:
            result['ФИО'] = sn + ' ' + gn
        else:
            result['ФИО'] = ''
        if key.decode('utf-8') == 'INN':
            result['ИНН'] = value.decode('utf-8')
        if key.decode('utf-8') == 'OGRN':
            result['ОГРН'] = value.decode('utf-8')
        if key.decode('utf-8') == 'SNILS':
            result['СНИЛС'] = value.decode('utf-8')
        if key.decode('utf-8') == 'O':
            result['Организация'] = value.decode('utf-8')
        if key.decode('utf-8') == 'L':
            l = value.decode('utf-8')
        if key.decode('utf-8') == 'street':
            street = value.decode('utf-8')
        if l and street:
            result['Адрес'] = l + ' ' + street

    return result
