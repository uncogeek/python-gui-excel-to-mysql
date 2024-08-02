# Employee Data Insertion GUI Application

This Python application provides a graphical user interface (GUI) to insert data from an Excel file into a MySQL database. The application is built using PyQt5.

## Features

- Simple and user-friendly GUI
- Reads data from an Excel file
- Inserts data into a MySQL database

## Requirements

- Python 3.x
- PyQt5
- pandas
- mysql-connector-python

## Installation

1. **Clone the repository:**

    ```bash
    git clone https://github.com/yourusername/employee-data-insertion-gui.git
    cd employee-data-insertion-gui
    ```

2. **Install the required Python packages:**

    ```bash
    pip install -r requirements.txt
    ```

3. **Create a MySQL database:**

    Create a database named `python` in your MySQL server.

4. **Import the SQL file:**

    Import the `edata.sql` file into the created database to set up the necessary table.

    ```bash
    mysql -u yourusername -p python < edata.sql
    ```

5. **Update database connection information:**

    Open `gui-employee.py` and update the MySQL database connection information with your credentials.

    ```python
    self.mydb = mysql.connector.connect(
        host="your_host",
        user="your_username",
        password="your_password",
        database="python"
    )
    ```

## Usage

1. **Run the application:**

    ```bash
    python gui-employee.py
    ```

2. **Using the application:**

    - Choose an Excel file from the `src` folder by clicking the "Choose File" button.
    - Click the "Add to DB" button to insert the data into the MySQL database.

## Contributing

Feel free to fork this repository and make changes. Pull requests are welcome.

## License

This project is licensed under the MIT License.

## Contact

For any questions or suggestions, please open an issue or contact me.
