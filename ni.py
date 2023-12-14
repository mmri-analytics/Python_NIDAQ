import nidaqmx
import matplotlib.pyplot as plt
import matplotlib.animation as animation
import pandas as pd
import time
import os
import numpy as np
from datetime import datetime
import glob

# ... [previous code] ...
# Directories
directory = os.getcwd() + "\\dt\\testCsv\\"
excel_directory = os.getcwd() + "\\dt\\testSummary\\"

# Create directory for CSV files if it doesn't exist
if not os.path.exists(directory):
    os.makedirs(directory)
    print(f"Directory '{directory}' created successfully.")
else:
    print(f"Directory '{directory}' already exists.")
    
if not os.path.exists(excel_directory):
    os.makedirs(excel_directory)
    print(f"Directory '{excel_directory}' created successfully.")
else:
    print(f"Directory '{excel_directory}' already exists.")
# Function to get the latest trial number
def get_latest_trial_number(directory):
    files = glob.glob(directory + '*.csv')
    if not files:
        return 0
    else:
        # Extract trial numbers from file names
        trial_numbers = [int(f.split('-trial')[1].split('.')[0]) for f in files if '-trial' in f]
        return max(trial_numbers) if trial_numbers else 0

# Get the latest trial number
trial_number = get_latest_trial_number(directory) + 1
# Constants
sampling_frequency = 5e3
mag_threshold = 0.0012
time_elapse = 4
duration = 25



# Function to calculate RMS
def calculate_rms(data):
    return np.sqrt(np.mean(np.square(data)))

# Initialize variables
wdt = []
excel_filename = excel_directory + 'signal_data.xlsx'

# Check if the Excel file exists, if not create it
if not os.path.exists(excel_filename):
    pd.DataFrame(columns=['Trial', 'RMS', 'Max', 'Min','Date']).to_excel(excel_filename, index=False)

# Start measuring
start_time = time.perf_counter()
with nidaqmx.Task() as task:
    datas = []
    measureTime = []
    segDats = []
    segTimeStamp = [start_time]

    # Add an analog input channel
    try:
        task.ai_channels.add_ai_voltage_chan("Dev1/ai0")
    except:
        task.ai_channels.add_ai_voltage_chan("cDAQ2Mod1/ai0")

    task.timing.cfg_samp_clk_timing(rate=sampling_frequency, sample_mode=nidaqmx.constants.AcquisitionType.CONTINUOUS, source="OnboardClock")
    last_trigger_countand=0
    while (time.perf_counter() - start_time) < duration:
        data = task.read(number_of_samples_per_channel=1)
        wdt.append(data[0])
        measureTime.append(time.perf_counter() - start_time)
        datas.append(data[0])

        if data[0] > mag_threshold:
            if segTimeStamp[0] == start_time:
                print(round(time.perf_counter() - segTimeStamp[0]))
                segTimeStamp[0] = time.perf_counter()
        current_time = time.perf_counter()
        elapsed_time = current_time - segTimeStamp[0]
        if (elapsed_time // time_elapse > last_trigger_countand) and (len(datas) > sampling_frequency * time_elapse * 0.9) and (segTimeStamp[0] != start_time):
            last_trigger_count = elapsed_time // time_elapse
            print(time.perf_counter() - segTimeStamp[0])
            spanL = round(-time_elapse * sampling_frequency)
            current_time_str = datetime.now().strftime('%Y-%m-%d-%H-%M-%S')
            filename = f"{directory}{current_time_str}-trial{trial_number}.csv"
            if spanL < 1.5 * time_elapse * sampling_frequency:
                # Create DataFrame with specified column names
                df = pd.DataFrame({'time': measureTime[spanL:], 'magnitude': datas[spanL:]}).to_csv(filename)
            else:
                # Create DataFrame with specified column names for the entire data
                df = pd.DataFrame({'time': measureTime, 'magnitude': datas}).to_csv(filename)
            datas = []
            measureTime = []

# Calculate RMS, Max, and Min and append to Excel file
rms_value = calculate_rms(wdt)
max_value = max(wdt)
min_value = min(wdt)
today_date = datetime.now().strftime('%Y-%m-%d')  # Get today's date

df = pd.read_excel(excel_filename)
new_row = {'Trial': trial_number, 'RMS': rms_value, 'Max': max_value, 'Min': min_value, 'Date': today_date}
df = df.append(new_row, ignore_index=True)
df.to_excel(excel_filename, index=False)



