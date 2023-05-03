# Excel to Word

This is a web application built using Flask that allows you to generate Word documents based on a Word template and data from an Excel file. With just a few clicks, you can produce customized Word documents that meet your specific needs. This app is perfect for creating invoices, contracts, reports, and other documents that require data from an Excel file to be integrated into a Word template.

## Getting Started

To get started, clone this repository and navigate to the directory in your terminal. Then, install the required packages by running the following command:

```bash
pip install -r requirements.txt
```

Once the packages are installed, you can run the app by executing the following command:

``` bash
python app.py
```

You can access the app in your browser by navigating to `http://localhost:8080/`.

## Usage

To use the app, simply upload your Excel file containing the relevant data, along with a Word template that includes placeholders formatted in the `{Key}` format. Once you have uploaded both files, click the "Generate" button, and voila! Your personalized Word document is ready for use. The app will automatically generate a URL where you can download the Word document.

## Built With

- Flask
- Google Cloud Storage
- OpenPyXL
- python-docx

## Live

The app is deployed on GCP and domain: [excel2word.app](https://excel2word.app)

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.
