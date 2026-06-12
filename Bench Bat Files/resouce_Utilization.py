# Import necessary libraries
import psutil  # For system usage data
import matplotlib.pyplot as plt  # For graphing
import time  # For time-related functions
import datetime  # For date and time manipulation
import os  # For file operations

# Define the constants
# Path used to store the data and graph
DATA_DIR = 'system_usage_data'
# Setting the time interval for 10 seconds
INTERVAL = 60
# Size of graph figure
FIGURE_SIZE = (10, 6)

def collect_data():

    memory_usage = []  # List to store memory usage data
    cpu_usage = []  # List to store CPU usage data
    time_data = []  # List to store time data

    while True:
        # Getting the current time
        current_time = datetime.datetime.now()
        time_data.append(current_time)

        # Getting the memory usage percentage
        memory_percent = psutil.virtual_memory().percent
        memory_usage.append(memory_percent)

        # Getting the CPU usage percentage
        cpu_percent = psutil.cpu_percent()
        cpu_usage.append(cpu_percent)

        yield time_data, memory_usage, cpu_usage  # Yielding the data or combining

def save_data(time_data, memory_usage, cpu_usage):

    # time_data list used to represent time of each data point
    # memory_usage list used to represent memory usage percentages
    # cpu_usage list the CPU usage percentages
    with open(os.path.join(DATA_DIR, 'data.txt'), 'w') as f:
        # Writing data into the file
        for i in range(len(time_data)):
            f.write(f'{time_data[i]}, {memory_usage[i]}, {cpu_usage[i]}\n')

def create_graph(time_data, memory_usage, cpu_usage):

    # Creating the graph for visualization.

    # Creating the figure with the specified size
    plt.figure(figsize=FIGURE_SIZE)

    # Creating subplots for memory and CPU usage
    plt.subplot(1, 2, 1)
    plt.plot(time_data, memory_usage)  # Plotting memory usage
    plt.title('Memory Usage')  # Setting title
    plt.xlabel('Time')  # Setting x-axis label
    plt.ylabel('Usage (%)')  # Setting y-axis label
    plt.xticks(rotation=90)  # Rotating x-axis ticks

    plt.subplot(1, 2, 2)
    plt.plot(time_data, cpu_usage)  # Plotting CPU usage
    plt.title('CPU Usage')  # Setting title
    plt.xlabel('Time')  # Setting x-axis label
    plt.ylabel('Usage (%)')  # Setting y-axis label
    plt.xticks(rotation=90)  # Rotating x-axis ticks

    plt.tight_layout()  # Adjusting layout
    plt.savefig(os.path.join(DATA_DIR, 'usage_graph.png'))  # Saving the graph
    plt.close()  # Closing the figure

def create_hourly_graph(time_data, memory_usage, cpu_usage):

    # Creating the graph for hourly data.

    # Initializing dictionaries to store hourly data
    hourly_memory_usage = {}
    hourly_cpu_usage = {}

    # Group data by hour
    for i in range(len(time_data)):
        hour = time_data[i].hour
        if hour not in hourly_memory_usage:
            hourly_memory_usage[hour] = []
            hourly_cpu_usage[hour] = []
        hourly_memory_usage[hour].append(memory_usage[i])
        hourly_cpu_usage[hour].append(cpu_usage[i])

    # Calculate hourly averages
    hourly_memory_usage_avg = {hour: sum(values) / len(values) for hour, values in hourly_memory_usage.items()}
    hourly_cpu_usage_avg = {hour: sum(values) / len(values) for hour, values in hourly_cpu_usage.items()}

    # Create a figure with the specified size
    plt.figure(figsize=FIGURE_SIZE)

    # Create subplots for memory and CPU usage
    plt.subplot(1, 2, 1)
    plt.plot(hourly_memory_usage_avg.keys(), hourly_memory_usage_avg.values())  # Plot memory usage
    plt.title('Hourly Memory Usage')  # Set title
    plt.xlabel('Hour')  # Set x-axis label
    plt.ylabel('Usage (%)')  # Set y-axis label

    plt.subplot(1, 2, 2)
    plt.plot(hourly_cpu_usage_avg.keys(), hourly_cpu_usage_avg.values())  # Plot CPU usage
    plt.title('Hourly CPU Usage')  # Setting title
    plt.xlabel('Hour')  # Setting x-axis label
    plt.ylabel('Usage (%)')  # Setting y-axis label

    plt.tight_layout()  # Adjusting layout
    plt.savefig(os.path.join(DATA_DIR, 'hourly_usage_graph.png'))  # Saving graph
    plt.close()  # Closing the figure

def main():
    # Creating data directory if it doesn't exist
    if not os.path.exists(DATA_DIR):
        os.makedirs(DATA_DIR)

    # Initializing data collector
    collector = collect_data()

    while True:
        # Collecting data
        time_data, memory_usage, cpu_usage = next(collector)

        # Saving data
        save_data(time_data, memory_usage, cpu_usage)

        # Creating graphs
        create_graph(time_data, memory_usage, cpu_usage)
        create_hourly_graph(time_data, memory_usage, cpu_usage)

        # Waiting for the next interval
        time.sleep(INTERVAL)

if __name__ == '__main__':
    try:
        main()
    except KeyboardInterrupt:
        pass
    finally:
        print('Data collection is being stopped.')
