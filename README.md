# Excel File Manipulation with OneDrive and Microsoft Graph API

This project demonstrates how to connect to a OneDrive account and update an Excel file using the Microsoft Graph API. The project includes a Django web application for reading and writing Excel files, as well as for authenticating and accessing OneDrive.

The reading and writing steps require manual authorisation actions. Those steps can be improved by automating them to make the application more user-friendly. 

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
    - Create a [credentials.json](http://_vscodecontentref_/1) file in the root directory of the project (where `manage.py` is located) and save your credentials and file path:
      ```json
      {
          "client_id": "your_client_id",
          "client_secret": "your_client_secret",
          "tenant_id": "your_tenant_id",
          "file_path": "your_file_path"  # Update this with your file path
      }
      ```

4. **Ensure Proper Authentication**:
    - **Register Your Application in Azure**:
      - Go to the [Azure portal](https://portal.azure.com/).
      - Navigate to `Azure Active Directory` > `App registrations` > `New registration`.
      - Register your application and note down the `client_id` and `tenant_id`.
    - **Configure API Permissions**:
      - In the Azure portal, go to your registered application.
      - Navigate to `API permissions` > `Add a permission`.
      - Select `Microsoft Graph` > `Delegated permissions`.
      - Add the `Files.ReadWrite.All` permission.
      - Grant admin consent for the permissions.
    - **Create a Client Secret**:
      - In the Azure portal, go to your registered application.
      - Navigate to `Certificates & secrets` > `New client secret`.
      - Note down the `client_secret`.

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
    - You will be prompted to visit a URL for authentication. Copy the URL from the terminal, open it in your browser, and authenticate.
    - After authentication, you will receive an authorization code. Copy this code and paste it back into the terminal.

2. **Update Excel File**:
    - Navigate to `http://localhost:8000/update_excel/` to update the value of a specific cell in the Excel file.
    - You will be prompted to visit a URL for authentication. Copy the URL from the terminal, open it in your browser, and authenticate.
    - After authentication, you will receive another authorization code. Copy this code and paste it back into the terminal.
    - The process will automatically navigate to the read Excel file with the updated cell value.

## Project Structure

- `hello/auth.py`: Contains the authentication logic using MSAL.
- `hello/views.py`: Contains the views for reading and updating the Excel file.
- `hello/templates/hello/read_excel.html`: Template for displaying the cell value.
- `hello/templates/hello/update_excel.html`: Template for updating the cell value.

## Resources

- [Microsoft Graph API Documentation](https://docs.microsoft.com/en-us/graph/overview)
- [Working with Excel in Microsoft Graph](https://docs.microsoft.com/en-us/graph/api/resources/excel?view=graph-rest-1.0)

## License

This project is licensed under the MIT License.