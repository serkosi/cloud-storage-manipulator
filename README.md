# Excel File Manipulation with OneDrive and Microsoft Graph API

This project demonstrates how to connect to a OneDrive account and update an Excel file using the Microsoft Graph API. The project includes Jupyter notebooks for reading and writing Excel files, as well as for authenticating and accessing OneDrive.


## Prerequisites

- Python 3.x
- Jupyter Notebook
- Required Python libraries: `pandas`, `openpyxl`, `msal`, `requests`, `xlwings`

## Setup

1. **Clone the repository**:
    ```sh
    git clone https://github.com/serkosi/cloud-storage-manipulator.git
    cd cloud-storage-manipulator-main
    ```

2. **Install the required Python libraries**:
    ```sh
    pip install pandas openpyxl msal requests xlwings
    ```

3. **Set up Microsoft Graph API**:
    - Register your app in the Azure portal to get the necessary credentials (client ID and secret).
    - Add the required API permissions (Files.ReadWrite.All, Sites.ReadWrite.All).
    - Save your credentials in [credentials.json](http://_vscodecontentref_/7):
      ```json
      {
          "client_id": "your_client_id",
          "client_secret": "your_client_secret"
      }
      ```

## Usage

### Reading and Writing Excel Files

The [trial.ipynb](http://_vscodecontentref_/8) notebook demonstrates how to read from and write to Excel files using `pandas` and `openpyxl`.

### Connecting to OneDrive and Updating Excel Files

The [run.ipynb](http://_vscodecontentref_/9) notebook demonstrates how to authenticate with OneDrive using the Microsoft Graph API and update an Excel file.

1. **Run the Jupyter notebook server**:
    ```sh
    jupyter notebook
    ```

2. **Open [run.ipynb](http://_vscodecontentref_/10)**:
    - Follow the steps to authenticate with OneDrive and get an access token.
    - Use the access token to access and update the Excel file in OneDrive.

## Resources

- [Microsoft Graph API Documentation](https://docs.microsoft.com/en-us/graph/overview)
- [Working with Excel in Microsoft Graph](https://docs.microsoft.com/en-us/graph/api/resources/excel?view=graph-rest-1.0)

## License

This project is licensed under the MIT License.
