#!/bin/env python
#----------------------------------------------------------------------------
# Name:         deckbot.py
# Purpose:      Application to Create Powerpoint Decks given a Company ID
#
# Author:       Drew Fulton
# Created:      April 2020

import argparse

import presenters
import models
import views

''' Launches the entire application
'''

parser = argparse.ArgumentParser()
parser.add_argument("--id", action='store',
	help="Enter the ID of the Company", type=str, required=False)
args = parser.parse_args()

if args.id:
	presenters.DeckbotPresenter(company_id=args.id, view=views.DeckbotCLI())
else:
	presenters.DeckbotPresenter(view=views.DeckbotCLI())