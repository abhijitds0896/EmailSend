
import pandas as pd
import win32com.client as win32
from html import escape
from config import PERSON_NAME, DUE_DATE

# --------------------------------------------------
# Configuration
# --------------------------------------------------

EXCEL_FILE = r"C:\Users\adas\OneDrive - Sopra Steria\Finops\KeyVault&RG 1.xlsx"

CC_EMAIL = "gilles.tosi@soprasteria.com"

SUBJECT_PREFIX = "Action Required | Azure Stack Key Vault Usage"

# --------------------------------------------------
# Read all Excel sheets
# --------------------------------------------------

sheets = ["RED", "YELLOW", "GREEN", "BLUE"]

df = pd.concat(
    [pd.read_excel(EXCEL_FILE, sheet_name=sheet) for sheet in sheets],
    ignore_index=True
)

# Remove duplicate resources, if any
df.drop_duplicates(
    subset=["NAME", "RESOURCE GROUP", "Email Address"],
    inplace=True
)

# Remove rows without email
df = df.dropna(subset=["Email Address"])

# --------------------------------------------------
# Outlook initialization
# --------------------------------------------------

outlook = win32.Dispatch("Outlook.Application")

# --------------------------------------------------
# Group by project team email
# --------------------------------------------------

for email, group in df.groupby("Email Address"):

    email = str(email).strip()

    # Project / subscription details
    project_name = group["SUBSCRIPTION"].iloc[0]

    # Prepare Key Vault resource table
    html_table = group[
        ["NAME", "TYPE", "RESOURCE GROUP",
         "LOCATION", "SUBSCRIPTION"]
    ].to_html(
        index=False,
        border=1,
        justify="left"
    )

    # Create Outlook email
    mail = outlook.CreateItem(0)

    mail.To = email
    mail.CC = CC_EMAIL

    mail.Subject = (
        f"{SUBJECT_PREFIX} | {project_name}"
    )

    # --------------------------------------------------
    # Email body
    # --------------------------------------------------

    mail.HTMLBody = f"""
    <html>
    <body style="font-family:Arial;font-size:10pt;">

    <p>Hello Team,</p>

    <p>
    I hope you are doing well.
    </p>

    <p>
    As part of our Azure Stack service,
    we are conducting a review of Key Vaults deployed
    across the projects to understand their current
    usage, dependencies, and business criticality.
    </p>
    <p>
    Please find below the list of Key Vault resources
    identified under your project for review.
    </p>

    <p><b>Project / Subscription:</b>
    {escape(str(project_name))}</p>

    {html_table}

    <br>

    <p><b>Kindly review the listed Key Vaults and
    provide the following information:</b></p>
    <ol>
        <li>
            <b>Consumption Method:</b>
            How is the Key Vault accessed or consumed?
            (Azure Portal/UI, API, application integration,
            automation/scripts, or other methods).
        </li>
        <br>
        <li>
            <b>Stored Data &amp; Use Case:</b>
            What type of information is stored in the
            Key Vault (e.g., secrets, certificates,
            encryption keys), and what is its business
            or technical purpose?
        </li>

        <br>

        <li>
            <b>Dependency &amp; Business Impact:</b>
            Is the project dependent on the Key Vault
            for its day-to-day operations? Could the
            project continue operating without it?
            What would be the business or technical
            impact if the Azure Stack Key Vault service
            were no longer available?
        </li>
    </ol>

   <p>
    Please provide your Azure Key Vault usage details on or before
    <b>{escape(str(DUE_DATE))}</b>. Failing this, the resources may be considered for cleanup of the Azure Key Vault.
   </p>
    <p>
    If you have any questions or require clarification,
    please feel free to reach out.
    </p>

    <p>Thank you for your cooperation.</p>

    <br>

    Regards,<br>
    <b>{escape(str(PERSON_NAME))}</b><br>
    DISIT Cloud Operations Team

    </body>
    </html>
    """

    # --------------------------------------------------
    # Preview / Send
    # --------------------------------------------------

    # Uncomment below to send directly
    mail.Send()

    # Preview email before sending
    #mail.Display()

    print(f"Email prepared for: {email} | {project_name}")

