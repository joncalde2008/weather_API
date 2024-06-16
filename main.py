import requests
import csv
import os
import matplotlib.pyplot as plt
import openpyxl
from openpyxl.drawing.image import Image
import numpy as np

# OpenWeatherMap API key
api_key = '2a5033a3bd7b53d4a30d24b9fd62bcc2'

# List of city IDs
city_ids = '524901,703448,2643743,5128581,5368361,4887398,4699066,5391811,5392171'

# URL for the group endpoint
url = f'https://api.openweathermap.org/data/2.5/group?id={city_ids}&appid={api_key}'

# Fetch data from API
try:
    response = requests.get(url)
    response.raise_for_status()
    data = response.json()
except requests.exceptions.RequestException as e:
    print(f"Error fetching data from OpenWeatherMap: {e}")
    exit()

# Extract relevant fields
csv_header = ['City', 'Temperature (°C)', 'Pressure (hPa)', 'Humidity (%)', 'Min Temperature (°C)', 'Max Temperature (°C)']
weather_data = []

for city in data['list']:
    city_name = city['name']
    temp = city['main']['temp'] - 273.15  # Convert from Kelvin to Celsius
    pressure = city['main']['pressure']
    humidity = city['main']['humidity']
    temp_min = city['main']['temp_min'] - 273.15  # Convert from Kelvin to Celsius
    temp_max = city['main']['temp_max'] - 273.15  # Convert from Kelvin to Celsius
    weather_data.append([city_name, temp, pressure, humidity, temp_min, temp_max])

# Calculate summary statistics
def calculate_statistics(data):
    return {
        'mean': np.mean(data),
        'median': np.median(data),
        'std_dev': np.std(data)
    }

temperature_stats = calculate_statistics([row[1] for row in weather_data])
pressure_stats = calculate_statistics([row[2] for row in weather_data])
humidity_stats = calculate_statistics([row[3] for row in weather_data])

summary_stats = [
    ['Mean', temperature_stats['mean'], pressure_stats['mean'], humidity_stats['mean'], '', ''],
    ['Median', temperature_stats['median'], pressure_stats['median'], humidity_stats['median'], '', ''],
    ['Standard Deviation', temperature_stats['std_dev'], pressure_stats['std_dev'], humidity_stats['std_dev'], '', '']
]

# Ensure the 'Excel' directory exists
if not os.path.exists('Excel'):
    os.makedirs('Excel')

# Write data to a CSV file
csv_file_path = 'Excel/data_weather.csv'
with open(csv_file_path, 'w', newline='', encoding='utf-8') as file:
    writer = csv.writer(file)
    writer.writerow(csv_header)
    writer.writerows(weather_data)
    writer.writerow([])  # Blank row for separation
    writer.writerows(summary_stats)

print(f'Weather data CSV file created at: {csv_file_path}')

# Plotting
city_names = [row[0] for row in weather_data]
temperatures = [row[1] for row in weather_data]
pressures = [row[2] for row in weather_data]
humidities = [row[3] for row in weather_data]

# Bar Chart
plt.figure(figsize=(12, 6))
plt.bar(city_names, temperatures, color='skyblue')
plt.xlabel('City')
plt.ylabel('Temperature (°C)')
plt.title('Temperature in Different Cities')
plt.xticks(rotation=45, ha='right')
plt.tight_layout()
bar_chart_path = 'Excel/bar_chart.png'
plt.savefig(bar_chart_path)
plt.close()

# Line Chart
plt.figure(figsize=(12, 6))
plt.plot(city_names, pressures, marker='o', color='green', linestyle='-')
plt.xlabel('City')
plt.ylabel('Pressure (hPa)')
plt.title('Pressure in Different Cities')
plt.xticks(rotation=45, ha='right')
plt.tight_layout()
line_chart_path = 'Excel/line_chart.png'
plt.savefig(line_chart_path)
plt.close()

# Scatter Chart
plt.figure(figsize=(12, 6))
plt.scatter(temperatures, humidities, color='red')
plt.xlabel('Temperature (°C)')
plt.ylabel('Humidity (%)')
plt.title('Temperature vs Humidity')
plt.tight_layout()
scatter_chart_path = 'Excel/scatter_chart.png'
plt.savefig(scatter_chart_path)
plt.close()

# Pie Chart
plt.figure(figsize=(8, 8))
plt.pie(temperatures, labels=city_names, autopct='%1.1f%%', startangle=140)
plt.title('Temperature Distribution')
plt.tight_layout()
pie_chart_path = 'Excel/pie_chart.png'
plt.savefig(pie_chart_path)
plt.close()

print(f'Charts saved in Excel directory.')

# Integrate charts into the workbook
# Path to save the final workbook
workbook_path = 'Excel/Data_and_graphs_weather.xlsx'

# Create a new workbook
workbook = openpyxl.Workbook()
sheet = workbook.active
sheet.title = 'data_weather'

# Read the CSV file and write it to the workbook
with open(csv_file_path, 'r', encoding='utf-8') as file:
    reader = csv.reader(file)
    for row in reader:
        sheet.append(row)

# Embed charts into the workbook
chart_files = [bar_chart_path, line_chart_path, scatter_chart_path, pie_chart_path]

row_offset = len(weather_data) + 4  # Starting row for the first chart after data and summary stats
for i, chart_path in enumerate(chart_files):
    try:
        img = Image(chart_path)
        img.anchor = f'A{row_offset}'
        sheet.add_image(img)
        row_offset += 20  # Adjust this value as needed for spacing between charts
    except Exception as e:
        print(f"Error embedding image {chart_path}: {e}")

# Save the workbook with embedded charts
try:
    workbook.save(workbook_path)
    print(f'Excel file with charts created at: {workbook_path}')
except Exception as e:
    print(f"Error saving workbook: {e}")

