@echo off
REM ============================================================
REM  test_azure_connection.bat
REM  Quickly verifies your Azure credentials and blob access
REM  before running the full service.
REM ============================================================
IF NOT EXIST ".venv\Scripts\activate.bat" (
    echo [ERROR] Run setup.bat first.
    pause & exit /b 1
)
call .venv\Scripts\activate.bat

echo.
echo  Testing Azure Blob Storage connection...
echo.
python -c "
import os, sys
from dotenv import load_dotenv
load_dotenv()

conn  = os.getenv('AZURE_CONNECTION_STRING','')
acct  = os.getenv('AZURE_ACCOUNT_NAME','')
key   = os.getenv('AZURE_ACCOUNT_KEY','')
cont  = os.getenv('AZURE_CONTAINER_NAME','powerbi-datasource')
csv_b = os.getenv('AZURE_CSV_BLOB_NAME','DataSetRawData_fromCSV.csv')
xl_b  = os.getenv('AZURE_EXCEL_BLOB_NAME','target.xlsx')

if not conn and not (acct and key):
    print('[FAIL] No Azure credentials found in .env')
    print('       Set AZURE_CONNECTION_STRING or AZURE_ACCOUNT_NAME + AZURE_ACCOUNT_KEY')
    sys.exit(1)

try:
    from azure.storage.blob import BlobServiceClient, StorageSharedKeyCredential
    if conn:
        svc = BlobServiceClient.from_connection_string(conn)
    else:
        cred = StorageSharedKeyCredential(acct, key)
        svc  = BlobServiceClient(account_url=f'https://{acct}.blob.core.windows.net', credential=cred)

    print(f'[OK]   Connected to storage account')

    cc = svc.get_container_client(cont)
    cc.get_container_properties()
    print(f'[OK]   Container found: {cont}')

    for blob_name in [csv_b, xl_b]:
        bc    = svc.get_blob_client(cont, blob_name)
        props = bc.get_blob_properties()
        size_gb = props.size / 1e9
        print(f'[OK]   Blob found: {blob_name}  ({size_gb:.2f} GB)')

    print()
    print('  All checks passed! Run run_local.bat to start the service.')

except Exception as e:
    print(f'[FAIL] {e}')
    sys.exit(1)
"
echo.
pause
