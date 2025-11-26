# Sheets Fantasy Football Platform v1.2

A tool to build, draft, and manage a fantasy football competition for one week's set of games or a subset of those within the Google Sheets platform using HTML popups and Google Apps Scripting.

Hey all, check out this newest version of a one-day/one-week fantasy football platform built in Google Sheets. The holidays with family or friends is a perfect time to bust this out to amp up the amusement of the 🦃Thanksgiving or 🎄Christmas games, a single Sunday, or the first round of the playoffs? 
What it does:

- 🏈 **Single Day (or week of the NFL) Leagues:** Run a full draft, competition, and recap for just a handful of games.
- 🚀 **Launch Setup Panels:** Instead of fumbling with cells, you launch a series of clean, web-app-like interfaces to guide you through setup.
- 👥 **Manage Group: **Add members, auto-generate fun team names (with alliteration!), and set the draft order by dragging-and-dropping or randomizing it with a dramatic, one-by-one reveal. You can even toggle on a 3rd Round Reversal draft. Randomly assign team names with a team name generator too.
- 🧮 **Customize Rosters:** Visually build your roster (e.g., QB, RB, RB, WR, FLX) and set the scoring format (1.0, 0.5, 0.0 PPR). You can create duplicate (clones) of players to ensure your whole gang can participate (you'll be prevented from drafting duplicates per individual roster
- 📆 **Select Matchups:** It pulls the latest weekly schedule from the ESPN API and lets you pick exactly which games you want to include in your player pool.
- 👌 **Make Picks:** Draft your team using the sheet's interface and checkboxes; you'll run your snake (or snake with 3RR) draft to fill out the roster, with prompts to confirm all picks and a penalty flag to recall a pick. Each player will have their Sleeper, ESPN, and Fantasy Pros projections, injury statuses, and previous score, when available.
- 🏆 **Get a Recap:** Once the games are over, it generates a summary screen showing the champion, podium finishers, top/bottom players at each position, "draft duds," and "undrafted gems."

To get rolling, first make a copy of the Sheets Fantasy Football Platform v1.2
The first menu option under "🏈 Football Tools" will prompt you to authorize the script. Click through the confirmations and then the menu should update to get you rolling with configuring your one-day or one-weekend


**🚀 INSTRUCTIONS:**
After authorizing the script, a new menu called "🏈 Football Tools" will
appear in your spreadsheet's menu bar. First you'll run the "▶️ Click Here to Setup"
Authorization will guide you through a few screens, one is tricky where you'll have 
to click the "Advanced" option in the lower left to proceed through the authorization.

Once you've authorized the script, the toolbar should reload (refresh if not)
You should be prompted to start working through the setup, which is designed to be
completed in order, with checkmarks (✔️) appearing as you complete each. Details below.

⚙️ USAGE - SETUP MENU:

1. 👥 Members & Draft Order
   - Add/remove members and assign them fun, auto-generated team names.
   - Set the draft order by dragging-and-dropping (Manual) or have the
     script randomize it for you upon submission.
   - Configure advanced draft settings like "3rd Round Reversal".

2. 📆 Game Selection
   - Automatically loads the current NFL week's schedule.
   - Select which NFL matchups you want to include in your contest's player pool.

3. 🧮 Roster Configuration
   - Define the structure of each team's roster (e.g., 1 QB, 2 RBs, etc.).
   - Set the league's scoring format (PPR: 1.0, 0.5, or 0.0).
   - The tool will automatically calculate player availability and warn you
     if you need to add "Extra Copies" of players to support your league size.

4. 🚀 Finalize & Deploy
   - Opens a final confirmation panel that summarizes all your settings.
   - From here, you can launch back into any of the setup panels to make last-minute
     changes or click "Submit Setup" to lock in your configuration and prepare for the draft.

I'm thrilled you're here and checking this out. If you're feeling generous and have the
means to support my work, you can support my wife, five kiddos (sixth on the way!), and me:
https://www.buymeacoffee.com/benpowers
OR
https://venmo.com/benpowerscreative

Thanks for checking out the script!
