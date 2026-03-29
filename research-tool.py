import tkinter
import tkinter.messagebox
import sqlite3
import json
import traceback
import sys
from datetime import datetime
import openai
from json_repair import repair_json
import os
from dotenv import load_dotenv

# Import libraries: tkinter for GUI, slite3 for db, json for GPT function call, traceback for
# better error handling, sys for program exit, openai for interaction with GPT, json_repair to handle GPT errors with
# json format

# Structure for GPT response function
GET_TOPIC_INFO_WITH_SUBTOPICS = {
    # Nested dictionary structure for GPT response
    "name": "store_topic_information_and_explore_subtopics",
    "description": "Store information about a topic and then explore more subtopics",
    "parameters": {
        "type": "object",
        "properties": {
            "information": {
                "type": "string",
                "description": "Information about the topic",
            },
            "subtopics": {
                "type": "array",
                "description": "List of subtopics to explore further",
                "items": {
                    "type": "object",
                    "description": "subtopic to explore further",
                    "properties": {
                        "name": {
                            "type": "string",
                            "description": "The name of the subtopic"
                        }
                    }
                }
            }
        }
    }
}
GET_TOPIC_INFO_WITHOUT_SUBTOPICS = {
    # Nested dictionary structure for GPT response
    "name": "store_topic_information",
    "description": "Store information about a topic",
    "parameters": {
        "type": "object",
        "properties": {
            "information": {
                "type": "string",
                "description": "Information about the topic",
            }
        }
    }
}

class Logging:
    def __init__(self, filename):
        self.__file = open(datetime.now().isoformat().replace(':', '_') + filename, "w")

    def log(self, log_message):
        self.__file.write(log_message + "\n")
        self.__file.flush()

    def close(self):
        self.__file.close()
# -----------------------------------------------------------------------------------------------------
# Class to manage sqlite database creation
class DatabaseManager:

    def __init__(self, logging):
        self.__conn = None
        self.__cur = None
        self.logging = logging

    def execute_sqlite_command(self, query, parameters=None):
        # Establish a connection to the SQLite database and execute the provided query
        self.logging.log(f'Executing SQL Command: {query}')
        with sqlite3.connect('topic_presentation.db') as self.__conn:
            self.__cur = self.__conn.cursor()
            if parameters:
                self.__cur.execute(query, parameters)
            else:
                self.__cur.execute(query)
            # Fetch all the results after executing the query
            output = self.__cur.fetchall()
        return output

    def drop_table(self):
        # Drop the 'topic_exploration' table if it already exists in the database
        self.logging.log("Dropping topic_exploration_table")
        self.execute_sqlite_command('DROP TABLE IF EXISTS topic_exploration')
        print('Table dropped from database if already exists')

    def create_table(self):
        # Create the 'topic_exploration' table if it does not exist in the database
        self.logging.log("Creating topic_exploration_table")
        query = '''
            CREATE TABLE IF NOT EXISTS topic_exploration(
                topic_id INTEGER PRIMARY KEY NOT NULL,
                topic_level INTEGER,
                topic_name TEXT,
                topic_description TEXT,
                parent_topic_id INTEGER REFERENCES topic_exploration(topic_id)
            )
        '''
        self.execute_sqlite_command(query)
        print('Table created in database')

    def get_topic_id_by_name(self, topic_name):
        # Retrieve the topic_id for a given topic_name from the 'topic_exploration' table
        self.logging.log(f"Getting ID for {topic_name}")
        query = '''
            SELECT * FROM topic_exploration WHERE topic_name = (?)'''
        topic_row = self.execute_sqlite_command(query, (topic_name,))
        topic_id = topic_row[0][0]
        return topic_id

    def get_topics_by_parent_id(self, parent_topic_id):
        # Retrieve topics based on given parent topic ID from the database
        query = '''
            SELECT * FROM topic_exploration WHERE parent_topic_id = ?'''
        result = self.execute_sqlite_command(query, (parent_topic_id,))
        return result

    def get_topic_by_id(self, topic_id):
        # Retrieve a topic based on given topic ID from the database
        query = '''
            SELECT * FROM topic_exploration WHERE topic_id = ?'''
        result = self.execute_sqlite_command(query, (topic_id,))
        return result[0]

    def get_main_topic_children(self):
        query = '''
            SELECT topic_id, topic_name FROM topic_exploration WHERE parent_topic_id = 1'''
        result = self.execute_sqlite_command(query)
        return result

    def get_list_of_subtopics(self):
        query = '''
            SELECT topic_name FROM topic_exploration'''
        result = self.execute_sqlite_command(query)
        # list comprehension
        print([i[0] for i in result])
        return [i[0] for i in result]

    def insert_row(self, topic_level, topic_name, topic_description, parent_topic_id):
        # Insert a new row into the 'topic_exploration' table
        self.logging.log(f"Adding {topic_name} to the database")
        query = '''
            INSERT INTO topic_exploration (topic_level, topic_name, topic_description, parent_topic_id)
            VALUES (?, ?, ?, ?)
        '''
        self.execute_sqlite_command(query, (topic_level, topic_name, topic_description, parent_topic_id))
        print('Topic added to database')

    def print_all_rows(self):
        # Display all rows from the 'topic_exploration' table
        query = 'SELECT * FROM topic_exploration'
        result = self.execute_sqlite_command(query)
        for row in result:
            print(row)


# -----------------------------------------------------------------------------------------------------
class Assistant:

    def __init__(self, logging):
        # Initialize attributes for the Assistant, including API key, system message, and response
        load_dotenv()
        self.__key = os.getenv("OPENAI_API_KEY")
        self.__openai_connection = openai.OpenAI(api_key=self.__key)
        self.logging = logging

    def get_client(self):
        # Get the API key for the Assistant
        return self.__openai_connection

    def get_key(self):
        return self.__key

    def ask_gpt_about_main_topic(self, main_topic, history):
        # Request info from GPT about the main topic
        self.logging.log(f"Getting information for main topic, \"{main_topic}\", from GPT")
        history.append({"role": "system",
                        "content": f"You are a research assistant. Help the user complete their research by giving detailed" 
                        f" information. Any responses or subtopics should  be highly relevant to '{main_topic}'"})

        functions = [GET_TOPIC_INFO_WITH_SUBTOPICS]
        function_call = {"name": GET_TOPIC_INFO_WITH_SUBTOPICS['name']}
        query_content = f"Tell me about \"{main_topic}\". List any relevant subtopics."
        return self.run_chat_completion_with_function_call(query_content, history, functions, function_call)

    def ask_gpt_about_sub_topic(self, current_topic, history, main_topic):
        # Request infor from GPT about the subtopic
        self.logging.log(f"Getting information for current topic, \"{current_topic}\", from GPT")
        functions = [GET_TOPIC_INFO_WITH_SUBTOPICS, GET_TOPIC_INFO_WITHOUT_SUBTOPICS]
        function_call = "auto"
        # Construct the user query and it to the conversation history
        query_content = (f"Tell me more about the subtopic '{current_topic}'.If there are any subtopics under"
            f" '{current_topic}' which aren't minor details our main topic, '{main_topic}', list those."
            f" Please omit anything we've already discussed.")
        return self.run_chat_completion_with_function_call(query_content, history, functions, function_call)

    def run_chat_completion_with_function_call(self, query_content, history, functions, function_call):
        client = self.get_client()
        print("Query:", query_content)
        query = {"role": "user", "content": query_content}
        history.append(query)

        # Send query to GPT model to get a response
        self.logging.log(f"Query: {query_content}")
        response = client.chat.completions.create(
            model="gpt-5.4-mini",
            messages=history,
            functions=functions,
            function_call=function_call,
            temperature=0.0,
            max_tokens=2000)
        print(response)
        self.logging.log(f"History of Conversation: {history}")

        # Extract JSON string from GPT response
        json_response_string = response.choices[0].message.function_call.arguments
        print(json_response_string)
        history.append(response.choices[0].message)
        return self.parse_gpt_response(json_response_string)

    def parse_gpt_response(self, json_response_string):
        try:
            # Try to parse JSON response
            gpt_response = json.loads(json_response_string)
        except json.decoder.JSONDecodeError:
            # Handle invalid JSON response by attempting repair
            print("ChatGPT returned invalid JSON, attempting repair")
            fixed_json_response_string = repair_json(json_response_string)
            print(fixed_json_response_string)
            try:
                # Try parsing the repaired JSON response
                gpt_response = json.loads(fixed_json_response_string)
            except:
                traceback.print_exc()
                print("Couldn't get a response from ChatGPT")
                sys.exit(1)
        # Update response attribute with GPT response and return it
        print('Response Returned')
        return gpt_response

    def send_summary_query(self, query_content):
        # Send a query to GPT for summarization
        client = self.get_client()
        print("GPT SUMMARY QUERY:", query_content)
        query = [{"role": "system", "content": "Please give a comprehensive summary of the provided content"},
                 {"role": "user", "content": query_content}]

        # Send query to GPT model to get a response
        response = client.chat.completions.create(
            model="gpt-5.4-mini",
            messages=query,
            temperature=0.0)
        print("GPT SUMMARY RESPONSE:")
        return response.choices[0].message.content


# -----------------------------------------------------------------------------------
class MyGUI:
    def __init__(self, assistant, logging):
        # Initialize the GUI with an Assistant instance
        self.assistant = assistant
        self.logging = logging
        self.topic_stack = []
        self.database = DatabaseManager(self.logging)
        # Drop and create the database table
        self.database.drop_table()
        self.database.create_table()

        # Create the main window, labels, buttons, entrybox, and textbox
        self.main_window = tkinter.Tk()
        self.main_window.geometry('1200x900')
        self.main_window.title('GPT Database Integration')
        self.create_frames()
        self.create_labels()
        self.create_buttons()
        self.create_entry_boxes()
        self.create_textbox()
        tkinter.mainloop()

    def create_frames(self):
        # Create frames for topics, responses, and buttons
        self.topic_frame = self.make_frame()
        self.topic_frame.pack()
        self.response_frame = self.make_frame()
        self.response_frame.pack()
        self.button_frame = self.make_frame()
        self.button_frame.pack()

    def create_labels(self):
        # Create labels for the user to input topics and display responses
        self.topic_label = self.make_label(self.topic_frame, 'Enter your query:', 1, 'ridge')
        self.topic_label.pack(side=tkinter.LEFT, ipadx=5, ipady=5, padx=20, pady=10)

        self.response_label = self.make_label(self.response_frame, 'Response:', 1, 'ridge')
        self.response_label.pack(ipadx=10, ipady=5, padx=20, pady=10)

    def create_buttons(self):
        # Create buttons for sending requests and quitting the program
        self.ok_button = self.make_button(self.button_frame, 'Send Request', self.get_user_input_and_start_gpt_dialog)
        self.ok_button.pack(padx=5, pady=10)
        self.quit_button = self.make_quit_button(self.button_frame)
        self.quit_button.pack(padx=5, pady=10)

    def create_entry_boxes(self):
        # Create an entry box for the user to input topics
        self.topic_entry = self.make_entry_box(self.topic_frame, 30)
        self.topic_entry.pack(side='left')

    def create_textbox(self):
        # Create a textbox for displaying responses
        self.response_textbox = self.make_textbox(self.response_frame)
        self.response_textbox.pack(side='left')

    # Methods to return tkinter widgets for the create methods
    def make_frame(self):
        return tkinter.Frame(self.main_window)

    def make_label(self, frame, text, border, relief):
        return tkinter.Label(frame, text=text,
                             border=border, relief=relief)

    def make_button(self, frame, text, command):
        return tkinter.Button(frame, text=text, command=command)

    def make_quit_button(self, frame):
        return self.make_button(frame, 'QUIT', self.quit_button_handler)

    def quit_button_handler(self):
        self.logging.close()
        self.main_window.destroy()

    def make_entry_box(self, frame, width):
        return tkinter.Entry(frame, width=width)

    def make_textbox(self, frame):
        return tkinter.Text(frame)

    def clear_displayed_data(self):
        self.response_textbox.delete('1.0', tkinter.END)

    def save_topic_to_db(self, topic_name, topic_info, topic_level, parent_topic_id):
        # Save main topic to database and retrieve its ID
        # Insert the main topic into the database
        self.database.insert_row(topic_level, topic_name, topic_info, parent_topic_id)
        # Get the ID for the main topic
        topic_id = self.database.get_topic_id_by_name(topic_name)
        self.database.print_all_rows()
        return topic_id

    def populate_topic_data(self, name, description, topic_level):
        # Display topic data from the database
        try:
            # Display the main topic in the response textbox
            self.response_textbox.insert(tkinter.END, '- ' * topic_level)
            self.response_textbox.insert(tkinter.END, name)
            self.response_textbox.insert(tkinter.END, ' > ')
            self.response_textbox.insert(tkinter.END, '\n')
            self.main_window.update()

        except Exception as e:
            # Show error message if an exception occurs while populating topic data
            error_message = f"An error has occurred: {str(e)}"
            tkinter.messagebox.showinfo('Error', error_message)
            traceback.print_exception(e)
            print(error_message)

    def recursive_get_topic_info(self, topic_name, history, parent_id, topic_level, main_topic):
        # Fetch topic info from GPT and save to database
        self.logging.log(f"Recursive method called for {topic_name}")
        if topic_level == 0:
            gpt_response = self.assistant.ask_gpt_about_main_topic(topic_name, history)
        else:
            gpt_response = self.assistant.ask_gpt_about_sub_topic(topic_name, history, main_topic)
        self.populate_topic_data(topic_name, gpt_response['information'], topic_level)
        topic_id = self.save_topic_to_db(topic_name, gpt_response['information'], topic_level, parent_id)
        if 'subtopics' in gpt_response:
            for sub_topic in gpt_response['subtopics']:
                sub_topic_name = sub_topic['name']
                self.recursive_get_topic_info(sub_topic['name'], history.copy(), topic_id, topic_level + 1, main_topic)

    def get_user_input_and_start_gpt_dialog(self):
        # Initiate the dialog with GPT and handle user input
        # Clear the textbox
        self.clear_displayed_data()
        # Get user input from the entry box
        user_input = str(self.topic_entry.get())
        # Check if the user input is not empty and run the recursive method to send info to GPT
        if user_input:
            self.recursive_get_topic_info(user_input, [], None, 0, user_input)
            self.clear_displayed_data()
            subtopic_list = self.database.get_list_of_subtopics()
            print(subtopic_list)
            output_data = self.summarize_topic()
            self.display_summary_data(output_data)
            subtopic_list = self.database.get_list_of_subtopics()
            delimiter = ','
            input_string = delimiter.join(subtopic_list)
            openai_api_key = self.assistant.get_key()
            print()
            print("POWER POINT CODE:")
            print()
            self.logging.log(f"Printing VBA PowerPoint Code")
            print(generate_vba_code(input_string, openai_api_key))



    def display_summary_data(self, output_data):
        # Displays retrieved data on the GUI
        self.logging.log(f"Displaying output data in the GUI text box")
        for item in output_data:
            self.response_textbox.insert(tkinter.END, item[0])
            self.response_textbox.insert(tkinter.END, '\n')
            self.response_textbox.insert(tkinter.END, item[1])
            self.response_textbox.insert(tkinter.END, '\n')
            self.main_window.update()

    def recursive_summarize_subtopic(self, topic_id):
        # Recursively fetches info about subtopics
        self.logging.log(f"Recursive method called for subtopics")
        current_list = []
        current_list.append(self.database.get_topic_by_id(topic_id))
        child_list = self.database.get_topics_by_parent_id(topic_id)
        for child in child_list:
            child_tree_info = self.recursive_summarize_subtopic(child[0])
            current_list.extend(child_tree_info)
        return current_list

    def get_gpt_summary_query(self, current_list):
        # Generates a query for GPT summarization based on topics list
        # Returns constructed query for GPT
        query = ""
        for topic in current_list:
            query += '\n'
            query += topic[2]  # Topic Name
            query += '\n'
            query += topic[3]  # Topic Description
        return query

    def summarize_topic(self):
        # Summarizes main topic children along with their subtopics
        starting_list = self.database.get_main_topic_children()
        output_data = []
        for item in starting_list:
            # Summarize subtopics for main topic, generate query based on details of subtopics,
            # send query to the assistant and retrieve summary, and return summarized data for main topics
            self.logging.log(f"Calling methods to get summary data from GPT")
            current_list = self.recursive_summarize_subtopic(item[0])
            query_content = self.get_gpt_summary_query(current_list)
            summary_content = self.assistant.send_summary_query(query_content)
            output_data.append((item[1], summary_content))
        return output_data

def generate_vba_code(input_string, openai_api_key):
    # Split the input string into lines
    lines = input_string.split(',')

    # Initialize the VBA code for PowerPoint
    vba_code = "Sub CreatePowerPointSlides()\n"
    vba_code += "Dim ppApp As Object\n"
    vba_code += "Dim ppPres As Object\n"
    vba_code += "Dim ppSlide As Object\n"
    vba_code += "Set ppApp = CreateObject(\"PowerPoint.Application\")\n"
    vba_code += "Set ppPres = ppApp.Presentations.Add\n"

    # Set up OpenAI API
    openai.api_key = openai_api_key

    # Generate VBA code for each line
    for line in lines:
        prompt = f"Create a VBA code snippet to add a new slide in PowerPoint with the text: \"{line}\""
        response = openai.completions.create(model="5.4-mini", prompt=prompt, max_tokens=150)
        vba_code += response.choices[0].text.strip() + "\n"

    # Close the VBA subroutine
    vba_code += "End Sub"

    return vba_code


if __name__ == '__main__':
    logging = Logging("program_log.log")
    assistant = Assistant(logging)
    mygui = MyGUI(assistant, logging)




 


