# LinkedIn post — WorkMate AI

Second project in my portfolio: **WorkMate AI**.

The idea started from a very simple problem: during a normal workday we constantly switch between Outlook, Teams, Excel, Power BI and other tools, repeating many of the same actions every morning.

So I built a small **personal work assistant** with **Python + Flask** that brings these workflows into one local dashboard.

WorkMate AI can currently:

- create a Morning Brief from recent Outlook emails
- identify unread messages and potential priorities
- prepare and send Outlook emails
- open Excel, Word, PowerPoint, Teams and Power BI
- prepare Teams messages
- show local notifications and alerts
- connect to an open Power BI Desktop semantic model
- read tables, columns and measures
- run DAX queries and reuse existing measures
- understand simple commands even with small typing errors thanks to fuzzy matching

The part I enjoyed most was the **Power BI Desktop integration**.

Through DAX Studio, WorkMate can inspect the semantic model and reuse the DAX measures that already exist in the report.

That means a request such as:

> “How many items are still pending?”

can be mapped to the correct measure in the model instead of duplicating the business logic.

Everything runs locally on Windows.

The goal is not to replace the tools I already use, but to **reduce repetitive context switching between them**.

I am using these projects to learn Python by starting from real problems and turning them into tools I can actually use.

GitHub:
https://github.com/AndreaJoeleMontone/workmate-ai

#Python #PowerBI #Automation #Flask #DAX #Microsoft365 #Productivity #GitHub #Programming
