import pandas as pd
import psycopg2
from psycopg2.extras import execute_values
import os
import boto3
from io import BytesIO

# Configuração do PostgreSQL
DB_HOST = os.getenv("PGHOST")
DB_PORT = os.getenv("PGPORT")
DB_USER = os.getenv("PGUSER")
DB_PASSWORD = os.getenv("PGPASSWORD")
DB_NAME = os.getenv("PGDATABASE")

# Configuração do S3
S3_BUCKET = os.getenv("BUCKET")
S3_REGION = os.getenv("REGION")
S3_ENDPOINT = os.getenv("ENDPOINT")
S3_ACCESS_KEY = os.getenv("ACCESS_KEY_ID")
S3_SECRET_KEY = os.getenv("SECRET_ACCESS_KEY")

def connect_db():
    try:
        conn = psycopg2.connect(
            host=DB_HOST,
            port=DB_PORT,
            user=DB_USER,
            password=DB_PASSWORD,
            database=DB_NAME
        )
        print("✓ Conectado ao PostgreSQL")
        return conn
    except Exception as e:
        print(f"✗ Erro ao conectar: {e}")
        return None

def get_s3_client():
    try:
        s3 = boto3.client(
            's3',
            endpoint_url=S3_ENDPOINT,
            aws_access_key_id=S3_ACCESS_KEY,
            aws_secret_access_key=S3_SECRET_KEY,
            region_name=S3_REGION
        )
        print("✓ Conectado ao S3")
        return s3
    except Exception as e:
        print(f"✗ Erro S3: {e}")
        return None

def import_table(df, conn, table_name, columns):
    cursor = conn.cursor()
    try:
        data = []
        for _, row in df.iterrows():
            data.append(tuple(row.get(col) for col in columns))
        
        cols_str = ", ".join(columns)
        execute_values(
            cursor,
            f"INSERT INTO {table_name} ({cols_str}) VALUES %s",
            data,
            template=None
        )
        conn.commit()
        print(f"✓ {len(data)} registros em '{table_name}'")
    except Exception as e:
        conn.rollback()
        print(f"✗ Erro em {table_name}: {e}")
    finally:
        cursor.close()

def main():
    print("=== Importador CSV ===\n")
    
    conn = connect_db()
    if not conn:
        return
    
    s3_client = get_s3_client()
    if not s3_client:
        conn.close()
        return
    
    try:
        response = s3_client.list_objects_v2(Bucket=S3_BUCKET)
        
        if 'Contents' not in response:
            print("✗ Nenhum arquivo no bucket")
            return
        
        csv_file = None
        for obj in response['Contents']:
            if obj['Key'].endswith(('.xlsx', '.xls', '.csv')):
                csv_file = obj['Key']
                break
        
        if not csv_file:
            print("✗ Nenhum arquivo encontrado")
            return
        
        print(f"Baixando: {csv_file}")
        response = s3_client.get_object(Bucket=S3_BUCKET, Key=csv_file)
        file_data = BytesIO(response['Body'].read())
        
        if csv_file.endswith(('.xlsx', '.xls')):
            xls = pd.ExcelFile(file_data)
            sheets = xls.sheet_names
        else:
            sheets = ['dados']
        
        mapping = {
            'disciplinas': ('disciplinas', ['Nome da disciplina', 'Serie', 'Curso', 'Turno']),
            'editores': ('editores', ['e-mail autorizado']),
            'grades_salvas': ('grades_salvas', ['Data', 'Curso', 'Serie', 'Conteudo', 'Turno', 'Turma']),
            'professores': ('professores', ['Id', 'Nome', 'Disciplinas', 'Disponibilidade']),
            'salas': ('salas', ['Curso', 'Série', 'Turno', 'Turma', 'Sala'])
        }
        
        for sheet in sheets:
            sheet_lower = sheet.lower().strip()
            
            if csv_file.endswith(('.xlsx', '.xls')):
                df = pd.read_excel(file_data, sheet_name=sheet)
            else:
                df = pd.read_csv(file_data)
            
            df.columns = df.columns.str.strip()
            
            print(f"\nProcessando: {sheet}")
            
            for key, (table, cols) in mapping.items():
                if key in sheet_lower:
                    import_table(df, conn, table, cols)
                    break
        
        print("\n✓ Concluído!")
        
    except Exception as e:
        print(f"✗ Erro: {e}")
    finally:
        conn.close()

if __name__ == "__main__":
    main()
