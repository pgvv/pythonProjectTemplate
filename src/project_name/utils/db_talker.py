# --------------------------- DOCUMENTATION ---------------------------#
# --------------------------- IMPORTS ---------------------------#
import os

import clickhouse_connect as ch
import mysql.connector
import pandas as pd
from dotenv import load_dotenv
from loguru import logger

# --------------------------- CONSTANTS ---------------------------#
# --------------------------- FUNCTIONS ---------------------------#
load_dotenv()


class BetaMarketPlace:
    def __init__(self) -> None:
        try:
            host = os.getenv("BMP_HOST")
            db = os.getenv("BMP_DB")
            user = os.getenv("BMP_USER")
            password = os.getenv("BMP_PASS")

            self.conn = mysql.connector.connect(
                host=host, user=user, password=password, database=db
            )
            self.csr = self.conn.cursor(dictionary=True)
        except Exception as e:
            logger.error(f"Error fetching DB creds: {e}")

    def test_connection(self):
        """Tests connection to BMP DB"""
        try:
            self.csr.execute("SELECT CURDATE()")
            logger.info("BMP DB connection successful")
        except Exception as e:
            logger.error(f"BMP DB connection failed: {e}")

    def exit_connection(self):
        self.csr.close()
        self.conn.close()

    def query(self, sql_file: str, params: dict | None = None):
        """Executes query from SQL file and exports result as DataFrame"""
        try:
            with open(sql_file, encoding="utf-8") as f:
                query = f.read()

                self.csr.execute(query) if params is None else self.csr.execute(query, params)
                result = self.csr.fetchall()
                logger.info(f"{sql_file.split('/')[-1]} successfully queried")
        except Exception as e:
            logger.error(e)

        return pd.DataFrame(result)


class Clickhouse:
    def __init__(self) -> None:
        try:
            host = os.getenv("CH_HOST")
            port = int(os.getenv("CH_PORT", ""))
            user = os.getenv("CH_USER")
            password = os.getenv("CH_PASS", "")

            self.client = ch.get_client(
                host=host, port=port, username=user, password=password, query_limit=0
            )
        except Exception as e:
            logger.error(f"Error fetching DB creds: {e}")

    def query(self, sql_file: str, params: dict | None = None):
        """Executes query from SQL file"""
        try:
            with open(sql_file, encoding="utf-8") as f:
                query = f.read()
                logger.info(f"{sql_file.split('/')[-1]} successfully queried")
        except Exception as e:
            logger.error(e)

        return self.client.query_df(query, params)


# --------------------------- MAIN CODE ---------------------------#
if __name__ == "__main__":
    pass
