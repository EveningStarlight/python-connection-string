from sqlalchemy import create_engine, text
import urllib

class SQL:
    fileNameHeader = "excel_report_name"
    connectionHeader = "semantic_model_id"

    def __init__(self):

        # Database connection details
        server = 'frs-production-sql.database.windows.net'
        database = 'frs-production-sqldb'
        driver = 'ODBC Driver 17 for SQL Server'
        port = 1433  # Default port
        table = '[meta].[excel_report_semantic_model]'

        # Create a connection string with access token
        params = urllib.parse.quote_plus(
            f"DRIVER={driver};SERVER={server},{port};DATABASE={database};"
            f"Authentication=ActiveDirectoryInteractive"
        )
        connection_string = f"mssql+pyodbc:///?odbc_connect={params}"

        # Create an engine and connect to the database
        engine = create_engine(connection_string)

        # Use the engine to execute a query
        with engine.connect() as conn:
            self.result = conn.execute(text(f"SELECT TOP 10 * FROM {table}"))
            self.columns = [column[0] for column in self.result.cursor.description]
            self.fileNameColumn = self.columns.index(SQL.fileNameHeader)
            self.connectionColumn = self.columns.index(SQL.connectionHeader)
            self.rows = self.result.fetchall()

    def getConnectionFromFileName(self, fileName):
        for row in self.rows:
            if row[self.fileNameColumn] == fileName:
                return row[self.connectionColumn]
            
    def print(self):
        print(', '.join(self.columns))
        for row in self.rows:
            print(row)