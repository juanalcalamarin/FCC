
# FCC.AF

## FCC.AF.CargaSociedades

Azure Function que se encarga del procesado de sociedades a partir de un fichero excel que se carga en el hub de auditor�a interna.
|Par�metros de entrada  | Descripci�n|
|--|--|
| ScheduleTriggerTime | Programaci�n de la AF. Se ejecutar� una vez al d�a por la noche pero en caso de actualizaci�n urgente ser� posible ejecutarla de forma manual [0 0 1 * * *] <br> Para programar el trigger consultar el siguiente enlace: <br> https://learn.microsoft.com/en-us/azure/azure-functions/functions-bindings-timer?tabs=python-v2%2Cin-process&pivots=programming-language-csharp|
| _CertName | Nombre del certificado. Este certificado caduca anualmente y habr� que actualizarlo para que el proceso pueda autenticarse y escribir en SharePoint |
| _CertURL | Url del Key Vault donde se encuentra el certificado |
| _TenantId | ID del tenant |
| _ClientId | Id de la App |
| _ClientSecret | Secret de la App |
| _HubUrl | Url del hub |
| _TenantUrl | Url del tenant de SharePoint |
| _TenantAdminUrl | Url de la admin. de SharePoint |
| _SitePattern | Patr�n de los sitios de auditor�a interna. Por ejemplo: **"/sites/AUDINT-XXXX-PRE"** |
| _MembersGroup | grupo de members de auditoria interna. Este grupo ser� com�n a todos los sitios de cada a�o (2022, 2023, ...) **"auditoriainterna@fcces.onmicrosoft.com"**|
