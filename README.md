[README.md](https://github.com/user-attachments/files/32154963/README.md)
# SMAJ TeachPlan — Gemini Free-Tier Version v2

This version uses the Google Gemini API and now includes required Teacher Name and Grade Level fields.

## What teachers do

1. Upload the Curriculum Map (.docx)
2. Upload the Unit Plan (.docx)
3. Enter Teacher Name
4. Enter Grade Level
5. Enter the lesson topic
6. Optionally enter session/date
7. Click GENERATE LESSON PLAN
8. Download the generated Word (.docx) file

## Automatic Word placement

- Grade Level is automatically inserted into the lesson plan header.
- Teacher Name is automatically inserted in the "Prepared by" section.
- These values come directly from the form, so the AI cannot accidentally change or omit them.

## GitHub update

Replace these files in your repository:

- `app.py`
- `requirements.txt` (same Gemini requirements; replacing it is safe)

Keep `.streamlit/config.toml`.

## Streamlit Secrets

GEMINI_API_KEY = "YOUR_REAL_GEMINI_API_KEY"
GEMINI_MODEL = "gemini-2.5-flash-lite"

Optional:
SITE_PASSWORD = "your-school-password"

Never upload your real API key to GitHub.


## Purple interface v3

The interface now uses a clean purple theme:
- purple buttons and accents
- soft lavender upload areas
- light purple input borders
- subtle purple page background
- teacher-friendly, uncluttered layout

For the full purple theme, update both `app.py` and `.streamlit/config.toml` on GitHub.


## Aesthetic Purple Interface v4

This version removes decorative icons and uses a cleaner, more elegant visual style:
- no icons or emojis in the interface
- soft purple and lavender palette
- rounded cards and input fields
- subtle shadows and gradients
- centered, minimalist title
- professional teacher-friendly appearance

Update both `app.py` and `.streamlit/config.toml` on GitHub.
