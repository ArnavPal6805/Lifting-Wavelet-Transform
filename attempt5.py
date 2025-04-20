import matlab.engine
import pandas as pd

# Start the MATLAB engine
eng = matlab.engine.start_matlab()

# Path to your MATLAB script
matlab_script = r'C:\Users\palar\OneDrive\Desktop\coding\Python\limited waveform transorm\lwt_decomp.m'

# Run the MATLAB script
eng.run(matlab_script, nargout=0)

# Read the results back from the Excel sheet created by MATLAB
filename = r'C:\Users\palar\OneDrive\Desktop\coding\Python\limited waveform transorm\Data_August_Renewable.xlsx'
result_df = pd.read_excel(filename, sheet_name='LWT Coefficients')

# Save the results to a new Excel file
output_file = 'Final_Output.xlsx'
result_df.to_excel(output_file, index=False)

print(f'Results successfully saved to: {output_file}')

# Close the MATLAB engine
eng.quit()
