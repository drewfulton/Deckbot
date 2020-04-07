#!/bin/env python
#----------------------------------------------------------------------------
# Name:         deckbot.py
# Purpose:      Application to Create Powerpoint Decks given a Company ID
#
# Author:       Drew Fulton
# Created:      April 2020


import presenters
import models
import views
# import interactors

''' Launches the entire application
'''

presenters.DeckbotPresenter(models.get_all_companies(offline=True), views.DeckbotCLI())