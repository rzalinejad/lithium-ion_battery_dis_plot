import os
import pandas as pd
import matplotlib.pyplot as plt

# Get the current directory (where your script and data files are located)
current_dir = os.path.dirname(os.path.abspath(__file__))

# Define the file names for each temperature
file_names = {
    '20°C': 'data_20°C.xlsx',  # Change these to the actual names of your Excel files
    '30°C': 'data_30°C.xlsx',
    '40°C': 'data_40°C.xlsx'
}

# Cut-off voltage
cutoff_voltage = 5.5  # Volts

# Create a dictionary to store data
data = {}

# Read and process each file
for temp, file_name in file_names.items():
    file_path = os.path.join(current_dir, file_name)
    df = pd.read_excel(file_path)
    
    # Extract columns
    time = df['Time[s]']
    voltage = df['Voltage[V]']
    current = df['Current[A]']  # Current is in Amps
    power = df['Power[W]']
    
    # Apply the cut-off voltage
    valid_data = df[voltage >= cutoff_voltage]
    
    # Check if valid data exists
    if valid_data.empty:
        print(f"Warning: No data above cut-off voltage for {temp}.")
        continue
    
    # Extract filtered values
    time = valid_data['Time[s]']
    voltage = valid_data['Voltage[V]']
    current = valid_data['Current[A]']
    power = valid_data['Power[W]']
    
    # Calculate time differences in hours (since time is in seconds)
    time_diff = time.diff().fillna(0) / 3600  # Convert seconds to hours
    
    # Calculate capacity in mAh
    capacity_mAh = (current * time_diff * 1000).cumsum()  # Convert to mAh
    
    # Calculate energy in Wh
    energy_Wh = (power * time_diff).cumsum()
    
    # Store data
    data[temp] = {
        'time': time,
        'voltage': voltage,
        'current': current,
        'power': power,
        'capacity': capacity_mAh,
        'energy': energy_Wh
    }

# Function to plot with maximum values and limited range
def plot_with_max(data, y_label, plot_type, title, unit):
    plt.figure(figsize=(12, 8))
    for temp, values in data.items():
        y_data = values[plot_type]
        max_value = y_data.max()
        max_time = values['time'][y_data.idxmax()]  # Get the time at which the max value occurs
        plt.plot(values['time'], y_data, label=f'{plot_type} at {temp}')
        plt.scatter(max_time, max_value, label=f'Max {temp} ({max_value:.2f} {unit})', marker='o')
    plt.title(title)
    plt.xlabel('Time [s]')
    plt.ylabel(y_label)
    plt.xlim([min(values['time']), max(values['time'])])  # Limit x-axis to valid range
    plt.ylim([min(y_data), max_value])  # Limit y-axis to valid range
    plt.grid(True)
    plt.legend()
    plt.show()

# Plot Time vs. Power with maximum values and cut-off range
plot_with_max(data, 'Power [W]', 'power', 'Power vs Time for Different Temperatures (Cut-off 5.5V)', 'W')

# Plot Time vs. Energy with maximum values and cut-off range
plot_with_max(data, 'Energy [Wh]', 'energy', 'Energy vs Time for Different Temperatures (Cut-off 5.5V)', 'Wh')

# Plot Time vs. Capacity with maximum values and cut-off range
plot_with_max(data, 'Capacity [mAh]', 'capacity', 'Capacity vs Time for Different Temperatures (Cut-off 5.5V)', 'mAh')

# Plot Time vs Voltage and Current with maximum values and cut-off range
plt.figure(figsize=(12, 8))
for temp, values in data.items():
    plt.plot(values['time'], values['voltage'], label=f'Voltage at {temp}', linestyle='--')
    plt.plot(values['time'], values['current'], label=f'Current at {temp}', linestyle='-')
    
    # Plot maximum values
    max_voltage = values['voltage'].max()
    max_current = values['current'].max()
    max_voltage_time = values['time'][values['voltage'].idxmax()]
    max_current_time = values['time'][values['current'].idxmax()]
    plt.scatter(max_voltage_time, max_voltage, label=f'Max Voltage {temp} ({max_voltage:.2f} V)', marker='o', color='blue')
    plt.scatter(max_current_time, max_current, label=f'Max Current {temp} ({max_current:.2f} A)', marker='o', color='orange')

plt.title('Voltage and Current vs Time for Different Temperatures (Cut-off 5.5V)')
plt.xlabel('Time [s]')
plt.ylabel('Voltage [V] / Current [A]')
plt.xlim([min(values['time']), max(values['time'])])  # Limit x-axis to valid range
plt.ylim([min(values['voltage']), max_voltage])  # Limit y-axis to valid voltage range
plt.grid(True)
plt.legend()
plt.show()

# Calculate and display total capacity and energy after cut-off
for temp, values in data.items():
    total_capacity = values['capacity'].iloc[-1]
    total_energy = values['energy'].iloc[-1]
    print(f"Total Capacity for {temp} after 5.5V cut-off: {total_capacity:.2f} mAh")
    print(f"Total Energy for {temp} after 5.5V cut-off: {total_energy:.2f} Wh")
