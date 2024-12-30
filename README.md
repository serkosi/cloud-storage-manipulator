# File Manipulation with OneDrive and Microsoft Graph API

This project demonstrates how to connect to a OneDrive account and update an Excel file using the Microsoft Graph API. The project includes a Django web application for reading and writing Excel files, as well as for authenticating and accessing OneDrive.

## Prerequisites

- Python 3.x
- Django
- Required Python libraries: `msal`, `requests`

## Setup

1. **Clone the repository**:
    ```sh
    git clone https://github.com/serkosi/cloud-storage-manipulator.git
    cd cloud-storage-manipulator-main
    ```

2. **Install the required Python libraries**:
    ```sh
    pip install django msal requests
    ```

3. **Set up Microsoft Graph API**:
    - Register your app in the Azure portal to get the necessary credentials (client ID and secret).
    - Add the required API permissions (Files.ReadWrite.All, Sites.ReadWrite.All).
    - Save your credentials in [credentials.json](http://_vscodecontentref_/1):
      ```json
      {
          "client_id": "your_client_id",
          "client_secret": "your_client_secret"
      }
      ```

## Running the Django Application

1. **Navigate to the Django template directory**:
    ```sh
    cd vscode-python-templates/django-template
    ```

2. **Apply migrations**:
    ```sh
    python manage.py migrate
    ```

3. **Run the development server**:
    ```sh
    python manage.py runserver
    ```

4. **Access the application**:
    Open your web browser and navigate to `http://localhost:8000`.

## Usage

### Reading and Updating Excel Files

1. **Read Excel File**:
    - Navigate to `http://localhost:8000/read_excel/` to read the value of a specific cell in the Excel file.

2. **Update Excel File**:
    - Navigate to `http://localhost:8000/update_excel/` to update the value of a specific cell in the Excel file.

## Project Structure

- [auth.py](http://_vscodecontentref_/2): Contains the authentication logic using MSAL.
- [views.py](http://_vscodecontentref_/3): Contains the views for reading and updating the Excel file.
- [read_excel.html](http://_vscodecontentref_/4): Template for displaying the cell value.
- [update_excel.html](http://_vscodecontentref_/5): Template for updating the cell value.

## Resources

- [Microsoft Graph API Documentation](https://docs.microsoft.com/en-us/graph/overview)
- [Working with Excel in Microsoft Graph](https://docs.microsoft.com/en-us/graph/api/resources/excel?view=graph-rest-1.0)

## License

This project is licensed under the MIT License.