"""
example_build_deck.py
Builds a demo deck that uses EVERY layout in humanized_ppt_tool.py.
Run:  python example_build_deck.py   ->  humanized_demo_deck.pptx
Replace the placeholder text/images with your own content.
Rule: facts, numbers and figures you pass in are placed verbatim - never altered.
"""
from PIL import Image, ImageDraw
from humanized_ppt_tool import Deck

# Placeholder figures (swap for your real chart/image files; they are inserted unmodified,
# scaled proportionally only).
def make_placeholder(path, w, h, label):
    im = Image.new("RGB", (w, h), (245, 247, 252))
    d = ImageDraw.Draw(im)
    d.rectangle([2, 2, w - 3, h - 3], outline=(160, 170, 195), width=3)
    d.line([60, h - 60, w - 60, 80], fill=(30, 39, 97), width=6)
    d.text((w // 2 - 80, h // 2), label, fill=(60, 60, 80))
    im.save(path)

make_placeholder("fig_square.png", 1000, 900, "Your chart here")
make_placeholder("fig_wide_1.png", 900, 650, "Figure A")
make_placeholder("fig_wide_2.png", 900, 650, "Figure B")

FOOT = "Team name | Project title"
d = Deck(palette="midnight")   # try: navy_orange, campus_blue, teal, forest, charcoal, terracotta, berry

d.add_title_roster_slide(
    title="Project Title Goes Here",
    subtitle="One sentence that says what the project found out and why it matters.",
    eyebrow="TOPIC \u2022 METHOD \u2022 METHOD",
    roster=[("Member One", "ID-001"), ("Member Two", "ID-002"), ("Member Three", "ID-003")],
    footer_note="Dataset: file_name.csv | 5,000 records | 20 variables")

d.add_stat_list_slide(
    title="Data Cleaning: From Raw File to Analysis-Ready Data", eyebrow="PART A",
    items=[("5,450 \u2192 5,200", "records retained after removing duplicates"),
           ("250", "duplicate or conflicting records removed"),
           ("115", "extreme values capped, not deleted")],
    note_heading="Cleaning log \u2014 what was fixed and how",
    note_text="Full log exported as cleaning_log.csv",
    footer=FOOT, page_num=2)

d.add_image_stats_slide(
    title="Exploratory Analysis: What the Data Shows", eyebrow="PART B",
    image_path="fig_square.png", image_caption="Correlation matrix",
    stats=[("39.2 yrs", "average age"), ("\u20b98.9L", "median income"),
           ("4.7 yrs", "average tenure"), ("7.5 / 10", "average satisfaction")],
    bullets=["First observation, written as one complete sentence with its number (r = 0.72).",
             "Second observation, same pattern: what moves with what, and what it means.",
             "Third observation."],
    footer=FOOT, page_num=3)

d.add_two_image_slide(
    title="Choosing the Number of Components", eyebrow="PART C",
    intro="The rule used, the number it gave, and the share of variance retained.",
    images=[("fig_wide_1.png", "Scree plot"), ("fig_wide_2.png", "Cumulative variance")],
    footer=FOOT, page_num=4)

d.add_grid_cards_slide(
    title="Four Segments \u2014 Profile and Names", eyebrow="PART C", cols=2,
    cards=[{"badge": "1", "heading": "Segment One", "stat": "948 (18.2%)",
            "lines": ["Fact line one", "Fact line two", "Fact line three"]},
           {"badge": "2", "heading": "Segment Two", "stat": "2,094 (40.3%)",
            "lines": ["Fact line one", "Fact line two", "Fact line three"]},
           {"badge": "3", "heading": "Segment Three", "stat": "1,056 (20.3%)",
            "lines": ["Fact line one", "Fact line two", "Fact line three"]},
           {"badge": "4", "heading": "Segment Four", "stat": "1,102 (21.2%)",
            "lines": ["Fact line one", "Fact line two", "Fact line three"]}],
    footer=FOOT, page_num=5)

d.add_bullets_image_slide(
    title="Cross-Checking the Result", eyebrow="EVIDENCE", intro="Why this matters",
    bullets=["Evidence point one.", "Evidence point two.", "Evidence point three."],
    image_path="fig_square.png", image_caption="Supporting figure",
    footer=FOOT, page_num=6)

d.add_stat_image_slide(
    title="A Simple Early-Warning Rule", eyebrow="INNOVATION",
    intro="One sentence that introduces the rule.",
    rule_text="Flag = 1  if  Score < 6  OR  Returns > 15%  OR  Complaints \u2265 4",
    stat_value="21.8%", stat_caption="of all customers are flagged today",
    note="That is over 1,100 customers to act on this quarter.",
    image_path="fig_wide_1.png", image_caption="Share flagged by segment",
    footer=FOOT, page_num=7)

d.add_process_flow_slide("From Data to Decision", [
    ("database", "Data", "Raw facts and metrics"),
    ("target", "Analysis", "Patterns emerge"),
    ("person", "Decision", "Bias-reduced choice")])

d.add_facts_slide(
    title="60-Second Verdict", eyebrow="CASE 03",
    paragraph="A short description of what this is, in two or three lines.",
    bullets=["How it works, step one", "Step two", "Step three"],
    facts=[("Format", "Individual"), ("Duration", "40 minutes"), ("Scoring", "Accuracy under time")])

d.add_tagged_list_slide("When Data Wins", "target", [
    ("Reliable data available", True), ("Patterns are stable", False),
    ("High cost of bias", True), ("Predictable environment", True)])

d.add_closing_bullets_slide(
    title="Using These Results Responsibly", eyebrow="RISKS AND CLOSE",
    bullets=["Limitation one.", "Limitation two.", "Limitation three."],
    summary_label="Bottom line", summary_text="One short paragraph a manager can act on.",
    footer=FOOT, page_num=11)

d.save("humanized_demo_deck.pptx")
print("saved humanized_demo_deck.pptx with", len(d.prs.slides._sldIdLst), "slides")
