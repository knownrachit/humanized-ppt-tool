# humanized-ppt-tool
Python tool that builds clean, natural-looking PowerPoint decks from structured content. It uses fixed palettes, 24pt body text, no AI-looking boxes and automatic text-fit. Works standalone or with Claude, ChatGPT or any chatbot that can run Python.
How someone can use it

Option 1: run it themselves by using the Humanized_PPT_Tool_Master_Prompt.md file in any ai chatbot that can run python.

step 1 - Open MASTER_PROMPT.md, paste Part 1 into Claude or any chatbot that can run Python, and give it the code (paste or upload).
step 2 - Send their topic, numbers and images. The bot builds the deck with the tool, checks the slides, and returns the .pptx.

Option 2 :
Install the requirements: pip install python-pptx pillow (Python 3.9+).
Clone the repo and run python example_build_deck.py. This builds an 11-slide demo that uses every layout.
Build their own deck:    
from humanized_ppt_tool import Deck
   d = Deck(palette="midnight")   # also: navy_orange, campus_blue, teal, forest, charcoal, terracotta, berry
   d.add_title_slide("My Title", "One-line subtitle")
   d.add_stat_slide("94%", "of tickets closed within SLA")
   d.save("my_deck.pptx")
   
