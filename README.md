# Hotel roster

A Streamlit prototype for creating and reviewing hotel staffing rosters.

## Try it locally

1. Install Python 3.11 or newer.
2. Create and activate a virtual environment.
3. Install the dependencies:

   ```bash
   pip install -r requirements.txt
   ```

4. Run the app:

   ```bash
   streamlit run app_roster.py
   ```

The app opens with anonymized sample employees. Use **Upload employee Excel** to use a team-specific workbook. The expected core columns are `id`, `name`, `role`, and `cross_train`; day-off columns are optional.

## Deploy with Streamlit Community Cloud

1. Create a GitHub repository and push this project.
2. At [share.streamlit.io](https://share.streamlit.io), select **Create app**.
3. Choose the repository and set the entry point to `app_roster.py`.
4. Deploy and share the generated `streamlit.app` link.

Every push to GitHub redeploys the app automatically.

## Data safety

Do not commit employee, payroll, preference, or exported roster files. They are deliberately excluded by `.gitignore`. Keep the repository private if the application code or the deployed app should only be available to your team.
