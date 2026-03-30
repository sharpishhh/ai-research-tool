# AI Research Tool

![Python](https://img.shields.io/badge/Python-3.12-blue)
![OpenAI](https://img.shields.io/badge/OpenAI-gpt--5.4--mini-brightgreen)
![Status](https://img.shields.io/badge/Status-Active-success)

![header](https://capsule-render.vercel.app/api?type=waving&color=gradient&height=150&section=header&text=AI%20Research%20Tool&fontSize=40&animation=fadeIn)

**An AI-powered desktop application that recursively explores a topic and its subtopics using the OpenAI API, stores results in a SQLite database, and displays a structured summary through a tkinter GUI.**

## Features

- Tkinter GUI for topic input and summary display
- OpenAI API integration using structured function calling for consistent JSON responses
- JSON repair fallback to handle malformed API responses gracefully
- Recursive topic exploration - automatically drills into subtopics until fully exhausted
- SQLite database storage of all topics, subtopics, and descriptions in a tree structure
- Summarization pass that compiles and displays results organized by main topic
- VBA code generation for PowerPoint slide creation based on explored topics
- Structured logging of all application events to a timestamped log file 

## Tech Stack

- Python 3.12
- Tkinter (GUI)
- OpenAI API (gpt-5.4-mini)
- SQLite3 (database)
- python-dotenv (environment variable managment)
- json-repair (malformed JSON handling)


## Setup

1. Clone the repo:
 `git clone https://github.com/sharpishhh/ai-research-tool.git`
2. Navigate into the project folder:
 `cd ai-research-tool`
3. Create and activate a virtual environment:
 `python -m venv venv`
 `source venv/Scripts/activate` (Windows/Git Bash)
 `source venv/bin/activate` (Mac/Linux)
4. Install dependencies:
 `pip instal -r requirements.txt`
5. Create an OpenAI account at https://platform.openai.com and generate an API key.
6. Create a `.env` file in the project root using `.env.example` as a template:
 `OPENAI_API_KEY=your_key_here`
7. Run the application:
 `python research-tool.py`
   
## Usage

Input a topic in the textbox next to "Enter your query:" and click the "Send Request" button. Summary data will begin appearing in the response box as the application recursively explores the topic until Chat GPT has exhausted it. A timestamped log file and SQLite database are generated automatically.VBA code for PowerPoint slide creation is printed to the console. Click "QUIT" to exit the application.

## Architecture

The application is organized into three classes:

**DatabaseManager** - handles all SQLite3 interactions including table creation, 
row insertion, and topic retrieval. It stores topics and subtopics in a 
self-referencing tree structure using parent topic IDs.

**Assistant** - manages all OpenAI API communication including structured function 
calling, JSON response parsing with a repair fallback, summarization queries, 
and VBA code generation.

**MyGUI** - constructs the tkinter interface and coordinates the application flow. 
It contains instances of both DatabaseManager and Assistant, and drives the core 
recursive logic that explores topics and subtopics until the API determines the 
subject is fully covered.

A standalone `Logging` class writes timestamped application events to a log file 
throughout the session.

## How It Works

1. **Launch:** Run the application and the tkinter GUI opens with an entry box and response area
2. **Input:** Enter a research topic and click "Send Request"
3. **Exploration:** The app calls the OpenAI API using structured function calling to retrieve 
   information and subtopics recursively until GPT determines the subject is fully covered
4. **Storage:** Each topic and subtopic is stored in a SQLite database in a tree structure 
   using parent topic IDs to preserve relationships
5. **Summarization:** Once exploration is complete, the app queries GPT to summarize each 
   main subtopic and its children
6. **Output:** Summaries are displayed in the GUI, all events are written to a timestamped 
   log file, and VBA code for PowerPoint slide generation is printed to the console
   
## Known limitations / Future improvements

### Known Limitations
- The Chat Completions API has been replaced by OpenAI's Responses API
- VBA is a legacy language exclusive to Microsoft Office. Output prints to the console 
  and requires a PowerPoint installation for use
- No maximum recursion depth — termination is determined by GPT, leaving a vulnerability 
  to excessive API calls and potential crashes
- `get_list_of_subtopics()` retrieves all topics including the main topic, not subtopics only
- Several debugging `print()` statements remain in production code

### Future Improvements
- Update Chat Completions API to Response API
- Swap VBA for python-pptx
- Rename `get_list_of_subtopics()` for accuracy
- Remove debugging `print()` statements
- Add depth limit as a safety net on recursion

## Credits

Originally developed as a university final project. Refactored and modernized in 2026 
to update outdated API integrations, migrate to environment variable credential 
management, and improve code organization.