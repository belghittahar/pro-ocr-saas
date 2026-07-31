# Professional OCR App

This Streamlit application uses the Google Cloud Vision API to perform Optical Character Recognition (OCR) on images. It supports dense text, pages, and complex layouts using `document_text_detection`.

## Setup Instructions

### 1. Create a Google Cloud Project and Enable Vision API

1. Go to the [Google Cloud Console](https://console.cloud.google.com/).
2. Create a new project or select an existing one.
3. Navigate to **APIs & Services > Library**.
4. Search for "Cloud Vision API" and click **Enable**.

### 2. Create a Service Account and Download the JSON Key

1. Navigate to **IAM & Admin > Service Accounts** in the Google Cloud Console.
2. Click **Create Service Account**, give it a name, and click **Create and Continue**.
3. Grant it the "Cloud Vision API User" role (or a similar role that grants access to the Vision API) and click **Continue**, then **Done**.
4. Click on the newly created service account email.
5. Go to the **Keys** tab.
6. Click **Add Key > Create new key**.
7. Choose **JSON** and click **Create**. The JSON file will download to your computer.
8. Rename this downloaded file to `google-credentials.json`.

### 3. Running Locally

1. Place the `google-credentials.json` file in the root directory of this project.
2. **Important:** Ensure `google-credentials.json` is added to your `.gitignore` file so you do not accidentally commit it to your repository.
3. Install the dependencies: `pip install -r requirements.txt`
4. Run the app: `streamlit run app.py`

### 4. Deploying securely on Streamlit Community Cloud

When deploying to Streamlit Community Cloud, you should **never** commit the `google-credentials.json` file to your GitHub repository. Instead, use Streamlit Secrets.

1. Go to your Streamlit Community Cloud dashboard and deploy your app from your GitHub repository.
2. Once deployed, go to the app's **Settings > Secrets**.
3. Open your downloaded `google-credentials.json` file in a text editor and copy its entire contents.
4. In the Streamlit Secrets text box, enter the following format:
   ```toml
   [gcp_service_account]
   type = "service_account"
   project_id = "your-project-id"
   private_key_id = "your-private-key-id"
   private_key = "your-private-key"
   client_email = "your-client-email"
   client_id = "your-client-id"
   auth_uri = "https://accounts.google.com/o/oauth2/auth"
   token_uri = "https://oauth2.googleapis.com/token"
   auth_provider_x509_cert_url = "https://www.googleapis.com/oauth2/v1/certs"
   client_x509_cert_url = "your-cert-url"
   ```
   (Alternatively, you can just paste the JSON content directly as a TOML dictionary if it aligns perfectly, but the standard Streamlit way is to convert the JSON key-value pairs to TOML under a header like `[gcp_service_account]`).

   **Easiest method for Secrets:**
   Streamlit supports reading nested dictionaries. Just add `[gcp_service_account]` on the first line and then paste the key-value pairs from your JSON file below it, ensuring they are formatted as TOML (e.g., replace `:` with `=`, remove trailing commas).
   
   Example:
   ```toml
   [gcp_service_account]
   type = "service_account"
   project_id = "..."
   private_key_id = "..."
   private_key = "-----BEGIN PRIVATE KEY-----\n...\n-----END PRIVATE KEY-----\n"
   client_email = "..."
   client_id = "..."
   auth_uri = "https://accounts.google.com/o/oauth2/auth"
   token_uri = "https://oauth2.googleapis.com/token"
   auth_provider_x509_cert_url = "https://www.googleapis.com/oauth2/v1/certs"
   client_x509_cert_url = "..."
   universe_domain = "googleapis.com"
   ```

5. Click **Save**. The app will now securely authenticate with Google Cloud without exposing your credentials.
