
# FCC.AF

## FCC.AF.CargaSociedades

Azure Function que se encarga del procesado de sociedades a partir de un fichero excel que se carga en el hub de auditoría interna.
|Parámetros de entrada  | Descripción|
|--|--|
| ScheduleTriggerTime | Programación de la AF. Se ejecutará una vez al día por la noche pero en caso de actualización urgente será posible ejecutarla de forma manual [0 0 1 * * *] <br> Para programar el trigger consultar el siguiente enlace: <br> https://learn.microsoft.com/en-us/azure/azure-functions/functions-bindings-timer?tabs=python-v2%2Cin-process&pivots=programming-language-csharp|
| _CertName | Nombre del certificado. Este certificado caduca anualmente y habrá que actualizarlo para que el proceso pueda autenticarse y escribir en SharePoint |
| _CertURL | Url del Key Vault donde se encuentra el certificado |
| _TenantId | ID del tenant |
| _ClientId | Id de la App |
| _ClientSecret | Secret de la App |
| _HubUrl | Url del hub |
| _TenantUrl | Url del tenant de SharePoint |
| _TenantAdminUrl | Url de la admin. de SharePoint |
| _SitePattern | Patrón de los sitios de auditoría interna. Por ejemplo: **"/sites/AUDINT-XXXX-PRE"** |
| _MembersGroup | grupo de members de auditoria interna. Este grupo será común a todos los sitios de cada año (2022, 2023, ...) **"auditoriainterna@fcces.onmicrosoft.com"**|
