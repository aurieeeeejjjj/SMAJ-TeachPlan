# Shareable Daily Lesson Plan Generator

This version is designed to be DEPLOYED ONLINE so your co-teachers can use one website link.

## What teachers see

They only need to:
1. Upload a Curriculum Map (.docx)
2. Upload a Unit Plan (.docx)
3. Enter the lesson topic
4. Optionally enter session/date
5. Click GENERATE LESSON PLAN
6. Download the generated Word (.docx) file

Teachers do NOT need to enter an OpenAI API key.

## Files

- `app.py` - website
- `requirements.txt` - Python packages
- `.streamlit/config.toml` - visual theme
- `.streamlit/secrets.toml.example` - example private settings

## Recommended simple deployment: Streamlit Community Cloud

1. Create a GitHub repository.
2. Upload these project files to that repository.
3. Do NOT upload a real API key.
4. In Streamlit Community Cloud, create a new app using the repository.
5. Set the main file path to:
   `app.py`
6. In the app's Secrets/settings area, add:

   OPENAI_API_KEY = "YOUR_REAL_KEY"
   OPENAI_MODEL = "gpt-5.6-luna"

7. Deploy the app.
8. Streamlit will give you a public website address that you can send to your co-teachers.

## Optional password

If you only want your co-teachers to use the site, add this to the private Secrets:

SITE_PASSWORD = "choose-a-password"

Then visitors must enter that password before they can use the generator.

## Important API cost note

The website owner's OpenAI API key pays for generations made through the website.
If many teachers use the site frequently, API usage can increase.

For shared use, `gpt-5.6-luna` is set as the default because it is designed for lower-cost,
high-volume workloads. You can change the model using the private `OPENAI_MODEL` secret.

## Security

Never type a real OpenAI API key directly into `app.py` and never publish it to GitHub.
Keep it only in the hosting service's private Secrets / environment-variable settings.
