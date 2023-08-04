# Ranking_automation
To run your Django project locally, you can follow the steps below:

## Prerequisites

- Python (version 3.x)
- Git 
- Virtualenv (optional but recommended)

## Local Setup

Follow the steps below to set up and run the project locally:

1. Clone the project repository:

   - git clone https://github.com/meissasoft/ranking_automation.git

3. Navigate to the project's root directory:

   - cd your-repository

5. (Optional) Create and activate a virtual environment to isolate the project's dependencies:
 
    - virtualenv venv # Create a virtual environment (optional)
    - source venv/bin/activate # Activate the virtual environment (Linux/Mac)
    - venv\Scripts\activate # Activate the virtual environment (Windows)

4. Install the project dependencies:

    - pip install -r requirements.txt

5. Start the cron job to run the bot:
   - python manage.py runcrons
