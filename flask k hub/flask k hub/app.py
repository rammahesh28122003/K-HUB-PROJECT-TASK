import pandas as pd
import matplotlib.pyplot as plt
from flask import Flask, render_template, request
import io
import base64
import pymongo

# Flask constructor
app = Flask(__name__)

# Connect to MongoDB
client = pymongo.MongoClient("mongodb://localhost:27017/")
db = client["mydatabase"]
data_collection = db["data_collection"]

# Root endpoint
@app.route('/')
def upload():
    return render_template('upload-excel.html')

@app.route('/view', methods=['POST'])
def view():
    # Read the File using Flask request
    file = request.files['file']
    # save file in the local directory
    file.save(file.filename)

    # Parse the data as a Pandas DataFrame type
    data = pd.read_excel(file)

    # Convert column names to strings explicitly
    data.columns = data.columns.astype(str)

    # Convert the DataFrame to a list of dictionaries
    data_list = data.to_dict(orient='records')
    # Insert the data into the MongoDB collection
    data_collection.insert_many(data_list)

    # Generate the pie chart
    labels = data.iloc[:, 0].tolist()  # Assuming the first column contains labels
    values = data.iloc[:, 1].tolist()  # Assuming the second column contains values

    # Generate the bar plot
    plt.figure(figsize=(10, 5))
    plt.bar(labels, values)
    plt.xlabel('Categories')
    plt.ylabel('Values')
    plt.title('Bar Plot')

    # Save the bar plot to a buffer
    bar_buffer = io.BytesIO()
    plt.savefig(bar_buffer, format='png')
    bar_buffer.seek(0)

    # Clear the figure to create a new one for the pie chart
    plt.clf()

    # Generate the pie chart
    plt.figure(figsize=(8, 8))
    plt.pie(values, labels=labels, autopct='%1.1f%%', startangle=140)
    plt.axis('equal')  # Equal aspect ratio ensures that pie is drawn as a circle.
    plt.title('Pie Chart')

    # Save the pie chart to a buffer
    pie_buffer = io.BytesIO()
    plt.savefig(pie_buffer, format='png')
    pie_buffer.seek(0)

    # Encode the binary image data to Base64
    pie_plot_base64 = base64.b64encode(pie_buffer.getvalue()).decode()
    bar_plot_base64 = base64.b64encode(bar_buffer.getvalue()).decode()

    # Return the HTML snippet that will render the table, pie chart, and bar plot
    return render_template('view-data.html', table=data.to_html(), pie_plot_base64=pie_plot_base64, bar_plot_base64=bar_plot_base64)


# Main Driver Function
if __name__ == '__main__':
    # Run the application on the local development server
    app.run(debug=True)
