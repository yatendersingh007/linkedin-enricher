# PASTE THIS INTO CLAUDE CODE (the Claude panel in VS Code)

Copy everything between the lines below and paste it as one message.

--------------------------------------------------------------------

I'm not a developer — explain simply and do the technical steps for me.
This folder is a Python "LinkedIn Enricher" project. It reads an Excel
file of event attendees containing LinkedIn profile URLs, uses the
TinyFish Agent API to extract each person's designation, company, and
location, and writes the results back into the Excel with checkpointing.

Please do the following, in order:

1. MOST IMPORTANT STEP: Install the TinyFish skill so you know how to
   use their API correctly:
   npx skills add github.com/tinyfish-io/tinyfish-cookbook --skill use-tinyfish
   Then check TinyFish's current documentation (https://docs.tinyfish.ai/
   — start from /llms.txt) and open enrich.py. The TinyFish client code
   in it (get_tinyfish_client and extract_profile) is a best guess —
   verify and fix the SDK import, client initialization, method call,
   and response parsing so it matches how the TinyFish Python SDK
   actually works today.

2. Check Python is installed (install if not) and install the
   dependencies from requirements.txt.

3. Create the .env file from .env.example and help me paste in my
   TinyFish API key when I have it ready.

4. I'll put my Excel file into the input/ folder. Look at it, tell me
   which column has the LinkedIn URLs, and set 'linkedin_url_column'
   in config.yaml to that exact header name.

5. Do a TEST RUN first on only 5 rows:
   python enrich.py --input input/<myfile>.xlsx --limit 5
   Show me the results in plain English and tell me roughly how many
   TinyFish steps/credits those 5 rows consumed, so I can estimate the
   cost of the full run before committing.

6. If the test looks good and I say go, run the full file. If it gets
   interrupted, resume it with:
   python enrich.py --input input/<myfile>.xlsx --resume

7. When done, tell me in plain English: how many rows came back OK,
   PARTIAL, FAILED, and SKIPPED, and where the enriched Excel is
   (it will be in output/).

If anything errors, diagnose and fix it yourself, then continue.
Never stop the whole run because one row failed — failed rows should
be marked FAILED in the TF Status column and skipped.

--------------------------------------------------------------------
