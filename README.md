# Deckbot MVP Overview
 
Goal: Given a CompanyID, produce a powerpoint deck that contains at least 3 slides (title slide, company overview, financial information using one metric).  This is modeled after Databook's FactPack.
	
## MVP Structure
- deckbot.py
  -	Main Application
  -	Main Application.
- models.py
  - Data management, pulls company and financial data from Databook API.
- views.py
  - Creates a CLI input for user to select a company from a list and sends Company object to Presenter for processing.
- presenters.py
  - Takes Company object from view, pulls data for that company from Model, and generates a powerpoint presentation using python-pptx.
	
## Dependencies
- requests
- python-pptx
