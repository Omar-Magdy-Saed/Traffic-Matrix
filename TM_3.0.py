import pandas as pd
import os
from openpyxl import load_workbook
from openpyxl.styles import Alignment
#import matplotlib.pyplot as plt
#import seaborn as sns
#import networkx as nx
#serviceimport matplotlib.pyplot as plt
import collections

# Welcome message

print (''' 
       ##################################################################################################
       ############################################### Hi, ##############################################
       # This is a simple tool to create a TM from extract CSV files , service and Trails , please note #
       # that there is still work in progress and will be updated with more features and options in the #
       #####  future.  for futher support please conntact me on omar.mostafa_mohamed@nokia.com  or ######
       #################          omar.mostafa.saed@gmail.com       #####################################                                                       
       ##################################################################################################
       ### Please Remember the full path of your input files including file name itself with extension ##
       ##################################################################################################
       ''')

# Define the functions

# Function to import the trails data
def import_trails_data():
    # Prompt the user for the file location
    trails_file = input("Please enter the location of the trails.csv file: ")

    # Load the CSV file into a DataFrame object
    trails_df = pd.read_csv(trails_file)

    # Keep only the specified columns in trails_df
    columns_to_keep_Trails = ['Rate', 'Name', 'Protection', 'From Node #1', 'From Port #1', 'To Node #1', 'To Port #1', 'Connection Type', 'Category', 'Frequency', '% Utilization', 'Channel Width', 'From NE #1', 'To NE #1']
    trails_df = trails_df[columns_to_keep_Trails]

    return trails_df

# Function to import the service data
def import_service_data():
    # Prompt the user for the file location
    service_file = input("Please enter the location of the service.csv file: ")

    # Load the CSV file into a DataFrame object
    service_df = pd.read_csv(service_file)

    # Keep only the specified columns in service_df
    columns_to_keep_Service = ['Rate', 'Name', 'Protection', 'From Node #1', 'From Port #1', 'To Node #1', 'To Port #1', 'Connection Type','Category', 'From NE #1', 'To NE #1' ]
    service_df = service_df[columns_to_keep_Service]

    return service_df

# Function to save the filtered DataFrame to an Excel file
def output_data(df, filename):
    # Check if the 'output' directory exists, if not, create it
    if not os.path.exists('output'):
        os.makedirs('output')

    # Save the DataFrame to an Excel file in the 'output' directory
    full_path = os.path.join('output', filename)
    df.to_excel(full_path, index=False)

    # Print the full path of the created file
    print(f"DataFrame saved to '{os.path.abspath(full_path)}'")
   
# Function to create a pivot table from the DataFrame and save it to an Excel file
def create_pivot_table(df, filename):

    # filter the dataframe to as per the rate type on the column 'Rate'
    # make a list of the unique values in the 'Rate' column and call it rates 
    rates = df['Rate'].unique()
    print ('the avilable rates on your network \n',rates)
    
    # Prompt the user to select a rate
    rate = input("Please enter the rate you want to filter by: ")

    # Filter the DataFrame to include only the selected rate
    filtered_df = df[df['Rate'] == rate]
    
    # Create a pivot table
    pivot_table = filtered_df.pivot_table(index='From NE #1', columns='To NE #1', aggfunc='size')

    # Save the pivot table to an Excel file in the 'output' directory
    full_path = os.path.join('output', filename)
    pivot_table.to_excel(full_path)

    # Open the Excel file and get the first worksheet
    book = load_workbook(full_path)
    sheet = book.active

    # Apply the formatting to the column headers and adjust the width of the columns
    for i, column in enumerate(sheet.columns):
        column = [cell for cell in column]
        if i == 0:  # first column
            max_length = 0
            for cell in column:
                try:  # Necessary to avoid error on empty cells
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            sheet.column_dimensions[column[0].column_letter].width = max_length
        else:  # other columns
            sheet.column_dimensions[column[0].column_letter].width = 4
        if column[0].row == 1:
            column[0].alignment = Alignment(text_rotation=90)
     
    # Rename the sheet with the selected rate
    sheet.title = rate

    # Save the changes and close the workbook
    book.save(full_path)
    book.close()

    # Print the full path of the created file
    print(f"Traffic Matrix for required rate saved to '{os.path.abspath(full_path)}'")

# Function to visualize the data
#def visualize_data(df, column1, column2):
#    # Histogram
#    plt.figure(figsize=(10, 6))
#    plt.hist(df[column1], bins=30, edgecolor='black')  # Customize the number of bins
#    plt.title(f'Histogram of {column1}')
#    plt.xlabel(column1)
#    plt.ylabel(column2)
#    plt.grid(True)
#    plt.show()

    # Boxplot
#    sns.boxplot(x=column2, y=column1, data=df)
#    plt.title(f'ploting of {column1} by {column2}')
#    plt.show()

def line_plot(df, column1, column2):
    plt.plot(df[column1], df[column2])
    plt.title(f'{column2} over {column1}')
    plt.xlabel(column1)
    plt.ylabel(column2)
    plt.show()

def density_plot(df, column):
    sns.kdeplot(df[column])
    plt.title(f'Density Plot of {column}')
    plt.xlabel(column)
    plt.ylabel('Density')
    plt.show()

def violin_plot(df, column1, column2):
    sns.violinplot(x=column1, y=column2, data=df)
    plt.title(f'Violin Plot of {column2} by {column1}')
    plt.show()

# Function to create a network graph
#def draw_connections_map(df):
#    # Create a new graph
#    G = nx.Graph()
#
#     # Add edges to the graph and count the number of trails between each pair of nodes
#    edge_counts = collections.Counter(zip(df['From Node #1'], df['To Node #1']))
#    for (node1, node2), count in edge_counts.items():
#        G.add_edge(node1, node2, count=count)
#
#    # Draw the graph
#    pos = nx.spring_layout(G)
#    #sizes = [500 * G.degree(node) for node in G.nodes]
#    nx.draw(G, pos, with_labels=True, node_size=1)
#
#    
#    # Draw edge labels and adjust edge thickness
#    edge_labels = nx.get_edge_attributes(G, 'count')
#    nx.draw_networkx_edge_labels(G, pos, edge_labels=edge_labels)
#    for (node1, node2, data) in G.edges(data=True):
#        width = data['count'] / 10
#        nx.draw_networkx_edges(G, pos, edgelist=[(node1, node2)], width=width)
#
#    plt.show()


###########################################################################
# Main Program
# Welcome message
print("Welcome to the Traffic Matrix Tool!")

# Prompt the user to select the type of TM
tm_type = input("Please enter the type of TM you want to create (Service or Trails): ")


# Select the appropriate DataFrame based on the user's response
if tm_type.lower() == 'service':
    
    # call the function and store the return values
    service_df = import_service_data()

    # call the function to save the filtered service_df to an Excel file in the 'output' directory
    output_data(service_df, 'filtered_service.xlsx')

    # Call the function to create a pivot table from service_df and save it to an Excel file
    create_pivot_table(service_df, 'TM_service.xlsx')

    # Call the function to visualize the data
    #visualize_data(service_df, 'Rate', 'Category')

    
    #call the function to density plot the data
    density_plot(service_df, 'Rate')
    #call the function to violin plot the data
    violin_plot(service_df, 'Rate', 'Category')

    

elif tm_type.lower() == 'trails':
    
    # call the function and store the return values
    trails_df = import_trails_data()

    # Call the function to save the filtered trails_df to an Excel file in the 'output' directory
    output_data(trails_df, 'filtered_trails.xlsx')

    # Call the function to create a pivot table from trails_df and save it to an Excel file
    create_pivot_table(trails_df, 'TM_trails.xlsx')

     # Call the function to visualize the data
    #visualize_data(trails_df, '% Utilization', 'Channel Width')

    #call the function to draw the connections map
    #draw_connections_map(trails_df)
    
else:
    print("Invalid input. Please enter either 'Service' or 'Trails'.")

    #wait for the user to close the program
    input("Press Enter to exit the program...")
    
    # Exit the program
    exit()

