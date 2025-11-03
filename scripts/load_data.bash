SAMPLE_DB=/mnt/c/git-codemie/ai_demo_plsql2java/db/samples/employees.sql.gz
DB_NAME=employees
PG_USER=demo
PG_PASSWORD=demo

# Set password environment variable to avoid password prompt
export PGPASSWORD=$PG_PASSWORD

# Use pg_restore for binary backup files
pg_restore --dbname=$DB_NAME --username=$PG_USER --host=localhost --port=5432 -Fc $SAMPLE_DB -c -v --no-owner --no-privileges
