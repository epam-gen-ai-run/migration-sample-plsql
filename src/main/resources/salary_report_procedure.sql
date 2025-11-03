-- Create a procedure that generates CSV file using pure PL/pgSQL
CREATE OR REPLACE FUNCTION generate_salary_report_to_csv(
    p_host VARCHAR DEFAULT 'localhost',
    p_port INTEGER DEFAULT 5432,
    p_database VARCHAR DEFAULT 'employees',
    p_output_file VARCHAR DEFAULT '/tmp/salary_report.csv'
)
RETURNS TEXT
LANGUAGE plpgsql
AS $$
DECLARE
    csv_content TEXT;
    rec RECORD;
BEGIN
    RAISE NOTICE 'Generating CSV file for database: % on %:%', p_database, p_host, p_port;
    
    -- Create CSV header
    csv_content := 'Department Name,Average Salary' || E'\n';
    
    -- Add data rows
    FOR rec IN 
        SELECT d.dept_name as dept_name, 
               ROUND(AVG(s.amount), 2) as avg_salary
        FROM employees.salary s
        JOIN employees.department_employee de ON s.employee_id = de.employee_id
        JOIN employees.department d ON de.department_id = d.id
        WHERE s.to_date > CURRENT_DATE AND de.to_date > CURRENT_DATE
        GROUP BY d.dept_name
        ORDER BY avg_salary DESC
        LIMIT 5
    LOOP
        csv_content := csv_content || 
                      '"' || REPLACE(rec.dept_name, '"', '""') || '",' || 
                      rec.avg_salary || E'\n';
    END LOOP;
    
    -- Try to export the CSV file
    BEGIN
        -- Use COPY to write the query results directly as CSV
        EXECUTE format('COPY (
            SELECT d.dept_name as "Department Name", 
                   ROUND(AVG(s.amount), 2) as "Average Salary"
            FROM employees.salary s
            JOIN employees.department_employee de ON s.employee_id = de.employee_id
            JOIN employees.department d ON de.department_id = d.id
            WHERE s.to_date > CURRENT_DATE AND de.to_date > CURRENT_DATE
            GROUP BY d.dept_name
            ORDER BY AVG(s.amount) DESC
            LIMIT 5
        ) TO %L WITH CSV HEADER', 
                      p_output_file);
        
        RETURN 'CSV file successfully created at: ' || p_output_file;
    EXCEPTION WHEN others THEN
        -- If COPY fails, return the CSV content for manual saving
        RETURN 'Could not write to file (requires superuser privileges). Here is the CSV content:' || E'\n' ||
               csv_content;
    END;
END;
$$;

-- Grant usage permissions
GRANT EXECUTE ON FUNCTION generate_salary_report_to_csv(VARCHAR, INTEGER, VARCHAR, VARCHAR) TO PUBLIC;

-- Example usage:
-- SELECT generate_salary_report_to_csv('localhost', 5432, 'employees', '/tmp/salary_report.csv');