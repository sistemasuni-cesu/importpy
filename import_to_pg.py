import pandas as pd
import psycopg2
from psycopg2.extras import execute_values
import os
from io import StringIO

from dotenv import load_dotenv
load_dotenv()

DATABASE_URL = os.getenv("DATABASE_URL")

def connect_db():
    try:
        conn = psycopg2.connect(DATABASE_URL)
        return conn
    except Exception as e:
        print(f"Erro ao conectar ao banco: {e}")
        return None

def import_disciplinas(df, conn):
    """Importa dados para tabela disciplinas"""
    cursor = conn.cursor()
    try:
        # Mapear colunas do CSV para a tabela
        data = []
        for _, row in df.iterrows():
            data.append((
                row.get('Nome da disciplina'),
                row.get('Serie'),
                row.get('Curso'),
                row.get('Turno')
            ))
        
        execute_values(
            cursor,
            "INSERT INTO disciplinas (nome_disciplina, serie, curso, turno) VALUES %s",
            data,
            template=None
        )
        conn.commit()
        print(f"✓ {len(data)} registros importados em 'disciplinas'")
    except Exception as e:
        conn.rollback()
        print(f"✗ Erro ao importar disciplinas: {e}")
    finally:
        cursor.close()

def import_editores(df, conn):
    """Importa dados para tabela editores"""
    cursor = conn.cursor()
    try:
        data = []
        for _, row in df.iterrows():
            data.append((row.get('e-mail autorizado'),))
        
        execute_values(
            cursor,
            "INSERT INTO editores (email_autorizado) VALUES %s",
            data,
            template=None
        )
        conn.commit()
        print(f"✓ {len(data)} registros importados em 'editores'")
    except Exception as e:
        conn.rollback()
        print(f"✗ Erro ao importar editores: {e}")
    finally:
        cursor.close()

def import_grades_salvas(df, conn):
    """Importa dados para tabela grades_salvas"""
    cursor = conn.cursor()
    try:
        data = []
        for _, row in df.iterrows():
            data.append((
                row.get('Data'),
                row.get('Curso'),
                row.get('Serie'),
                row.get('Conteudo'),
                row.get('Turno'),
                row.get('Turma')
            ))
        
        execute_values(
            cursor,
            "INSERT INTO grades_salvas (data, curso, serie, conteudo, turno, turma) VALUES %s",
            data,
            template=None
        )
        conn.commit()
        print(f"✓ {len(data)} registros importados em 'grades_salvas'")
    except Exception as e:
        conn.rollback()
        print(f"✗ Erro ao importar grades_salvas: {e}")
    finally:
        cursor.close()

def import_professores(df, conn):
    """Importa dados para tabela professores"""
    cursor = conn.cursor()
    try:
        data = []
        for _, row in df.iterrows():
            data.append((
                row.get('Id'),
                row.get('Nome'),
                row.get('Disciplinas'),
                row.get('Disponibilidade')
            ))
        
        execute_values(
            cursor,
            "INSERT INTO professores (id, nome, disciplinas, disponibilidade) VALUES %s",
            data,
            template=None
        )
        conn.commit()
        print(f"✓ {len(data)} registros importados em 'professores'")
    except Exception as e:
        conn.rollback()
        print(f"✗ Erro ao importar professores: {e}")
    finally:
        cursor.close()

def import_salas(df, conn):
    """Importa dados para tabela salas"""
    cursor = conn.cursor()
    try:
        data = []
        for _, row in df.iterrows():
            data.append((
                row.get('Curso'),
                row.get('Série'),
                row.get('Turno'),
                row.get('Turma'),
                row.get('Sala')
            ))
        
        execute_values(
            cursor,
            "INSERT INTO salas (curso, serie, turno, turma, sala) VALUES %s",
            data,
            template=None
        )
        conn.commit()
        print(f"✓ {len(data)} registros importados em 'salas'")
    except Exception as e:
        conn.rollback()
        print(f"✗ Erro ao importar salas: {e}")
    finally:
        cursor.close()

def main():
    """Função principal"""
    # Caminho do arquivo CSV (ajuste conforme necessário)
    csv_file = "horários aulas teste 3.xlsx"  # ou .csv
    
    if not os.path.exists(csv_file):
        print(f"✗ Arquivo não encontrado: {csv_file}")
        return
    
    # Conectar ao banco
    conn = connect_db()
    if not conn:
        return
    
    try:
        # Ler arquivo Excel/CSV com múltiplas abas
        print(f"Lendo arquivo: {csv_file}")
        
        if csv_file.endswith('.xlsx'):
            xls = pd.ExcelFile(csv_file)
            sheet_names = xls.sheet_names
            print(f"Abas encontradas: {sheet_names}")
        else:
            # Se for CSV, assume uma única aba
            sheet_names = ['dados']
        
        # Mapear abas para funções de importação
        import_functions = {
            'disciplinas': import_disciplinas,
            'editores': import_editores,
            'grades_salvas': import_grades_salvas,
            'professores': import_professores,
            'salas': import_salas
        }
        
        # Importar cada aba
        for sheet in sheet_names:
            sheet_lower = sheet.lower().strip()
            
            if csv_file.endswith('.xlsx'):
                df = pd.read_excel(csv_file, sheet_name=sheet)
            else:
                df = pd.read_csv(csv_file)
            
            # Limpar nomes de colunas (remover espaços extras)
            df.columns = df.columns.str.strip()
            
            print(f"\nProcessando aba: {sheet}")
            print(f"Colunas: {list(df.columns)}")
            
            # Encontrar função correspondente
            for key, func in import_functions.items():
                if key in sheet_lower:
                    func(df, conn)
                    break
        
        print("\n✓ Importação concluída!")
        
    except Exception as e:
        print(f"✗ Erro geral: {e}")
    finally:
        conn.close()

if __name__ == "__main__":
    main()
