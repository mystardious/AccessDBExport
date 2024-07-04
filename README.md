# Database Access App

A simple C# application to export queries from an Access database to an Excel file.

## Quick Usage Instructions

1. **Clone the Repository**
    ```sh
    git clone https://github.com/mystardious/AccessDBExport.git
    cd DatabaseAccessApp
    ```

2. **Build the Application**
    Open the project in your preferred C# IDE (such as Visual Studio) and build the solution.

3. **Run the Application**
    Use the following command to run the application:

    ```sh
    AccessDBExport --db "<DatabaseLocation>" --query "<QueryName>" --export "<ExportLocation>"
    ```

    - `--db`: Path to the Access database file.
    - `--query`: Name of the query to export.
    - `--export`: Path to export the query result to an Excel file.

4. **Example**
    ```sh
    AccessDBExport --db "c:\Users\Mystardious\Downloads\Reliability Reporting Database 24-25.mdb" --query "Cat Cause" --export "c:\Users\Mystardious\Downloads\export.xlsx"
    ```

## Help
For additional options and help, use the `--help` flag:
```sh
AccessDBExport --help
