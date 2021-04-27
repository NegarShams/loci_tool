"""
#######################################################################################################################
###											PSSE G74 Fault Studies													###
###		Script sets up PSSE to carry out fault studies in line with requirements of ENA G74							###
###																													###
###		Code developed by David Mills (david.mills@PSCconsulting.com, +44 7899 984158) as part of PSC 		 		###
###		project JK7938 - SHEPD - studies and automation																###
###																													###
#######################################################################################################################
"""

import os
import re
import datetime
import pandas as pd
from collections import OrderedDict

# Meta Data
__version__ = '0.1.0'

# Set to True to run in debug mode and therefore collect all output to window
DEBUG_MODE = True

# TODO: Define as a constant input
convert_to_kA = True
class plotting:
	"""
			General constants
		"""
	legend=['grid loci','windfarm R&X','windfarm flipped R&X','worst R&X point for V','worst R&X point for I']
	x_label='R (ohm)'
	y_lable='X (ohm)'
	axis_space = 5
	axis_step=20
	save_plot=False
	show_plot=False
	folder='plots'
	string_added='fig_'
	figure_type='.png'


class folder_file_names:
	dig_result_folder = 'dig_results'

class Digsilent:

	#object_name = 'POI 230kV FirstEnergy'
	object_name = 'POI 138 kV-ACE'
	#object_name = 'POI 230kV FirstEnergy'
	param_name_list= ['R, Re(Z) in Ohm', 'X, Im(Z) in Ohm']
	#param_name_list = ['Network Resistance in Ohm', 'Network Reactance in Ohm']
	param_number_list = [4, 5]
	number_or_name_toggle=True  # if false it means that it uses parameter name list to find the columns from
	# digsilent results, if True it uses number of columns
	header_list = ['H', 'R', 'X']

class General:
	"""
		General constants
	"""


	params_dict = dict()
	scalable_GSP_list = list()
	years_list = list()
	demand_scaling_list = list()
	station_dict = dict()
	loads_complete = bool()
	xl_file_name = str()

	loads_complete_f_str = 'False - Click to view loads summary'
	loads_complete_t_str = 'True'

	ext_csv = '.csv'
	node_label = 'Node Details'
	bus_name = 'Name'
	bus_voltage = 'Nominal (kV)'
	pre_fault = 'Pre-fault Voltage (p.u.)'
	bus_number = 'Busbar Number'
	x_r = 'X/R'

	# Default file types used for import / export
	file_types = (('xlsx files', '*.xlsx'), ('All Files', '*.*'))
	sav_types = (('PSSE (SAV) files', '*.sav'), ('All Files', '*.*'))
	# Nominal frequency assumed for studies
	nominal_frequency = 50.0

	def __init__(self):
		"""
			Just to avoid error message
		"""
		pass


class XlFileConstants:

	dill_file_name = 'station_dict.pkl'
	xl_checks_file_name = 'Load Checks Summary.xlsx'
	params_folder = 'params'

	excel_ws_name = 'MASTER Based on SubstationLoad'
	sheet1 = 'Complete Load Data'
	sheet2 = 'Missing Load Data'
	sheet3 = 'Amended Load Data'

	# define good data and bad data dataframes
	good_data = pd.DataFrame()
	bad_data = pd.DataFrame()
	#amend_data=pd.DataFrame() #suggested data in place of missing data

	# define columns from spreadsheet
	gsp_col_no = 0
	nrn_col_no = 1
	name_col_no = 2

	# define BPS and primary number of rows
	bsp_no_rows = 4
	prim_no_rows = 3

	# the number of rows between the end of a GSP and the next GSP (not including row with 'Average Cold Spell (ACS))
	row_separation = 1

	# define station type string
	gsp_type = 'GSP'
	bsp_type = 'BSP'
	primary_type = 'PRIMARY'
	pf_str = 'p.f'
	growth_rate_dict = dict()

	# define the column ranges of interest
	peak_mw_col = 7
	growth_rate_col = 10
	load_forecast_col_range = range(11, 25)
	seasonal_percent_col_range = range(26, 29)
	psse_buses_col_range = range(29, 37)

	# define cell on interest for pf
	pf_cell_tuple = (3, 7)

	def __init__(self):
		pass

 # class amend:
	#  estimate=False
 #
	#  def __init__(self):
	# 	 """
	#  		Just to avoid error message
	#  	"""
	# 	 pass

class SavedParamsStrings:

	params_file_name = 'SSE_LE_params.pkl'

	scalable_GSP_list_str = 'scalable_GSP_list'
	years_list_str = 'years_list'
	demand_scaling_list_str = 'demand_scaling_list'
	station_dict_str = 'station_dict'
	loads_complete_str = 'loads_complete'
	xl_file_name = 'xl_file_name'

	def __init__(self):
		"""
			Just to avoid error message
		"""
		pass


class Zones:
	zone_num = 'NUMBER'
	zone_name = 'ZONENAME'
	load_p = 'PLOAD'
	load_q = 'QLOAD'
	gen_p = 'PGEN'
	gen_q = 'QGEN'
	gen_p_controlled = 'PGEN_CONTROLLED'
	gen_p_other = 'PGEN_OTHER'
	gen_q_controlled = 'QGEN_CONTROLLED'
	gen_q_other = 'QGEN_OTHER'

	def __init__(self):
		pass


class GUI:
	"""
		Constants for the user interface
	"""
	# year_list = [datetime.datetime.now().year, datetime.datetime.now().year + 1, datetime.datetime.now().year+2]
	# season_list = ['Maximum Demand', 'Minimum Demand']
	# station_dict = dict()

	gui_name = 'PSC Load Estimates Tool'
	# 0.00 and 0.01 removed since these fault times will be added anyway
	default_fault_times = '0.06'

	# Default on whether the SAV case should be reloaded at the end of the fault
	# study or start from empty
	reload_sav_case = 1

	# Open excel with completed files
	open_excel = 1

	# Number of characters to fit into entry box for busbars
	busbar_box_size = 9

	# Number of busbar entry boxes on each row
	busbar_columns = 3
	vertical_busbars = 10
	empty_busbars = 40

	# Indicating the local directory
	local_directory = os.path.dirname(os.path.realpath(__file__))
	img_pth_main = os.path.join(local_directory, 'PSC Logo RGB Vertical.png')
	img_pth_window = os.path.join(local_directory, 'PSC Logo no tag-1200.gif')
	img_size = (128, 128)

	# Test to include on the GUI
	psc_uk = 'PSC UK'
	psc_phone = '\nPSC UK:  +44 1926 675 851'
	psc_font = 'Calibri 10 bold'
	psc_color_web_blue = '#%02x%02x%02x' % (43, 112, 170)
	psc_color_grey = '#%02x%02x%02x' % (89, 89, 89)

	load_radio_opts = OrderedDict([
		(0, 'None'),
		(1, 'All Loads'),
		(2, 'Selected GSP Only'),
		(3, 'Selected Zones Only')
	])

	gen_radio_opts = OrderedDict([
		(0, 'None'),
		(1, 'All Generators'),
		(2, 'Selected Zones Only')
	])

	def __init__(self):
		"""
			Purely to avoid error message
		"""
		pass


class GUIDefaultValues:
	"""
		Class will contain all the default values that are controlled / edited by the GUI
		and are therefore all stored in one location even though referenced from multiple
	"""
	# Thermal violation parameters:
	thermal_ol_pct = 100
	branch_rate = 'C'
	tx_rate = 'A'
	rpf_rate = 'C'

	# Bus voltage violation parameter, default is +/- 6% (i.e 0.94 p.u to 1.06 p.u):
	volt_range = [0.94, 1.06]

	# Voltage step violation parameters:
	v_step_limit = 0.03

	# Type of generation to scale
	scale_existing = True
	scale_end = True

	# Default name for GUI
	gui_name = 'Virtual Statcom and Capacity Evaluation Tool'

	# Default study is Virtual Statcom
	virtual_statcom_study = 1
	capacity_evaluation_study = 2

	# Whether generators should be added at end buses
	gen_end_buses = True
	gen_existing_buses = True

	# Whether to save the temporary folders or not
	temp_folders_save = True

	# Default hosting capacity is generation
	generation_hosting_capacity = 1
	load_hosting_capacity = 2

	# Case voltage defaults - Note can only contain 2 options due to method used in GUI for selecting case voltage
	case_voltage_options = {
		1: 11,
		2: 33
	}
	# Default option is 33kV
	default_case_voltage_option = 2

	# File types
	sav_types = (('PSSE (SAV) files', '*.sav'), ('All Files', '*.*'))
	xlsx_types = (('xlsx files', '*.xlsx'), ('All Files', '*.*'))

	# Default voltage parameters in the following order (
	# - lower voltage limit (p.u.)
	# - target voltage (p.u.)
	# - upper voltage limit (p.u.)
	# - voltage step limit (%)
	# - thermal_loading_limit (%)
	# default_limits = (
	# 	WpdThreshold.voltage_lower,
	# 	WpdThreshold.target_voltage,
	# 	WpdThreshold.voltage_upper,
	# 	WpdThreshold.step_change_limit * 100.0,
	# 	WpdThreshold.thermal_loading_limit * 100.0
	# )

	# Default contingency option for GUI radio button selection
	method_traditional = 'Traditional Planning'
	method_per_contingency = 'Per Contingency'
	method_ts_per_contingency = 'Time Series Per Contingency'
	study_method = {
		method_traditional: 1,
		method_per_contingency: 2,
		method_ts_per_contingency: 3
	}
	# Default method
	method_default = study_method[method_per_contingency]

	contingency_all = 'All Network Configurations'
	contingency_intact = 'Intact System Configuration Only'
	contingency_user_specified = 'Specific User Selected Contingencies'
	# Values used to identify relevant radio buttons
	contingency_values = {
		contingency_all: 1,
		contingency_intact: 2,
		contingency_user_specified: 3
	}
	# Default selected contingency
	contingency_default = contingency_values[contingency_intact]

	# Indicating the local directory
	local_directory = os.path.dirname(os.path.realpath(__file__))
	user_guide = 'JK7938-02 PSSE Load Estimate Importing Tool - Work Instruction.pdf'
	img_pth_psc_main = os.path.join(local_directory, 'PSC Logo RGB Vertical.png')
	img_pth_psc_window = os.path.join(local_directory, 'PSC Logo no tag-1200.gif')
	img_pth_wpd_main = os.path.join(local_directory, 'WPD Midlands English.jpg')
	hyperlink_psc_website = 'https://www.pscconsulting.com/'
	hyperlink_wpd_website = 'https://www.westernpower.co.uk/'
	img_size_psc = (120, 120)
	img_size_wpd = (150, 150)

	# Reference to the Tkinter binding of a mouse button
	mouse_button_1 = '<Button - 1>'

	font_family = 'Helvetica'
	psc_uk = 'PSC UK'
	psc_phone = '\nPSC UK:  +44 1926 675 851'
	psc_font = ('Calibri', '10', 'bold')
	psc_color_web_blue = '#%02x%02x%02x' % (43, 112, 170)
	psc_color_grey = '#%02x%02x%02x' % (89, 89, 89)

	# Set to WPD green
	color_main_window = '#%02x%02x%02x' % (222, 235, 247)
	# color_main_window = '#%02x%02x%02x' % (0, 176, 80)
	color_frame = 'white'
	color_scrollbar = 'light grey'
	color_pop_up_window = 'light grey'

	# Case voltage defaults - Note can only contain 2 options due to method used in GUI for selecting case voltage
	network_selection_options = {
		1: 'Radial BSP Networks',
		2: 'Parallel BSP Networks'
	}

	parallel_zone_str = 'zone'
	parallel_zone_name_str = 'zone name'

	def __init__(self):
		pass


class PSSE:
	"""
		Class to hold all of the constants associated with PSSE initialisation
	"""
	# Base MVA value assumed
	base_mva = 100.0

	# Maximum number of iterations for a Newton Raphson load flow (default = 20)
	max_iterations = 100
	# Tolerance for mismatch in MW/Mvar (default = 0.1)
	mw_mvar_tolerance = 1.0

	sid = 1

	# Load Flow Constants
	tie_line_flows = 0  # Don't enable tie line flows
	phase_shifting = 0  # Phase shifting adjustment disabled
	dc_tap_adjustment = 0  # DC tap adjustment disabled
	var_limits = 0  # Apply VAR limits immediately
	non_divergent = 0

	ext_bkd = '.bkd'

	# Default parameters for PSSE outputs
	# 1 = physical units
	def_short_circuit_units = 1
	# 1 = polar coordinates
	def_short_circuit_coordinates = 1

	# Minimum fault time that can be considered (in seconds)
	min_fault_time = 0.0001



	# Dependant on whether running in PSSE 33 or 34
	if 'PROGRAMFILES(X86)' in os.environ:
		program_files_directory = r'C:\Program Files (x86)\PTI'
	else:
		program_files_directory = r'C:\Program Files\PTI'

	# PSSE version dependant paths
	psse_paths = {
		32: 'PSSE32\PSSBIN',
		33: 'PSSE33\PSSBIN',
		34: 'PSSE34\PSSPY27'
	}
	os_paths = {
		32: 'PSSE32\PSSBIN',
		33: 'PSSE33\PSSBIN',
		34: 'PSSE34\PSSBIN'
	}

	# Relevant file names for PSSPY and PSSE when needing to search for them
	psspy_to_find = "psspy.pyc"
	pssarrays_to_find = "pssarrays.pyc"
	psse_to_find = "psse.bat"
	default_install_directory = r'C:\ProgramData\Microsoft\AppV\Client\Integration'

	# Default destination for PSSE output
	output_default = 1
	output_file = 2
	output_none = 6

	# Setting on whether PSSE should output results based on whether operating in DEBUG_MODE or not
	output = {True: output_default, False: output_none}

	def __init__(self):
		"""
			Purely to avoid error message
		"""
		pass


class BkdyFileOutput:
	"""
		Constants for processing BKDY file output
	"""
	base_mva = PSSE.base_mva

	# TODO: Add this a constants value
	if convert_to_kA:
		num_to_kA = 1000.0
		current_unit = 'kA'
	else:
		num_to_kA = 1.0
		current_unit = 'A'

	start = 'FAULTED BUS'
	current = 'FAULT CURRENT'
	impedance = 'THEVENIN IMPEDANCE'

	ik11 = "Ik'' ({})".format(current_unit)
	ip = 'Ip ({})'.format(current_unit)
	# Sum of DC components contributing to bus determine peak make
	ip_method1 = 'Ip sum of DC({})'.format(current_unit)
	# Peak calculated using x/r of thevenin impedance
	ip_method2 = 'Ip X/R method ({})'.format(current_unit)
	ibsym = 'Ibsym ({})'.format(current_unit)
	ibasym = 'Ibasym ({})'.format(current_unit)
	# Sum of DC components contributing to bus determine peak make
	ibasym_method1 = 'Ibasym ({})'.format(current_unit)
	# DC component calculated using x/r of thevenin impedance
	ibasym_method2 = 'Ibasym ({})'.format(current_unit)
	idc = 'DC ({})'.format(current_unit)
	# Sum of DC components contributing to bus determine peak make
	idc_method1 = 'DC from sum of DC({})'.format(current_unit)
	# DC component calculated from X/R at point of fault
	idc_method2 = 'DC X/R method({})'.format(current_unit)
	idc0 = 'DC_t0 ({})'.format(current_unit)
	v_prefault = 'V Pre-fault (p.u.)'

	# Impedance values
	x = 'X (p.u. on {:.0f} MVA)'.format(base_mva)
	r = 'R (p.u. on {:.0f} MVA)'.format(base_mva)

	# Error flag if Vpk returns infinity
	infinity_error = '*******'

	# Regex search expression broken down as follows:
	# (Infin)|(ity)|(\*{9} = Picks up the values returned if there is an error
	# (\*{9}) = Matches a 9 character * string which is returned for infinite values at time 0
	# (-{0,1}\d\.\d{4,5}(?!\d+\.)) =	Matches an optional - symbol followed by 4 or 5 numerical values where
	# 							there are not more numerical values and a decimal point following that point.
	# 							This will pick up the R and X values as well as the pre-fault voltage the optional -
	# 							allowing the values to be returned negative if the exist for error reporting.
	# (\d{1,3}\.\d{2}) = Matches for either a 1 to 3 decimal number followed by a decimal point and a 2 decimal number.
	# #					This will pick up angles.
	# (\d+\.\d) = Matches for any number of numerical values leading a decimal point with a single numerical value
	# 			afterwards.  This will pick up the fault current magnitudes.
	# #reg_search = re.compile('(\*{9})|(\d\.\d{4,5}(?!\d+\.))|(\d{1,3}\.\d{2})|(\d+\.\d)')
	reg_search = re.compile('(Infin)|(ity)|(\*{9})|(-?\d\.\d{4,5}(?!\d+\.))|(\d{1,3}\.\d{2})|(\d+\.\d)')
	# The following terms are used to confirm whether there are values returned which relate to an infinite value and
	# handled correctly.
	# TODO: May need to add an additional check to confirm that no values are returned as infinity when they shouldn't be
	nan_term1 = 'Infin'
	nan_term2 = 'ity'
	nan_term3 = '*' * 9

	# NaN value that is returned if error calculating fault current values
	nan_value = 'NaN'
	# This is replaced with the following and an error message given to user
	# TODO: Ensure error message is given to user
	nan_replacement = '0.0'

	def __init__(self):
		"""
			Purely to avoid error message
		"""
		pass

	def col_positions(self, line_type):
		"""
			Returns a dictionary with the associated column positions depending on the line type
		:param str line_type:  based on the values defined above returns the relevant column numbers
		:return dict, int (cols, expected_length):  Dictionary of column positions, expected length of list of floats
		"""
		# TODO: For peak fault current in make calculation should maximum of both methods be used
		cols = dict()
		if line_type == self.current:
			cols[self.ik11] = 0
			cols[self.ibsym] = 2
			# # Values no longer obtained from here since these relate to the values obtained by the sum of the
			# # calculated DC values rather than thevenin impedance as required by G74
			# #cols[self.idc] = 4
			# #cols[self.ibasym] = 5
			cols[self.idc_method1] = 4
			cols[self.ibasym_method1] = 5
			cols[self.ip_method1] = 6

			# Expected length of this list of floats
			expected_length = 7
		elif line_type == self.impedance:
			cols[self.r] = 0
			cols[self.x] = 1
			cols[self.v_prefault] = 2
			# Obtaining the DC, asym and peak values from the second row (THEVENIN ROW) is used since this
			# aligns with the requirements of the G74 standard rather to use the thevenin impedance
			cols[self.idc_method2] = 4
			cols[self.ibasym_method2] = 5
			cols[self.ip_method2] = 6
			# Not possible to export this data since in some cases get a result returned which says infinity
			# #cols[self.idc0] = 4
			# #cols[self.ibasym0] = 5
			# #cols[self.ip0] = 6

			# Expected length of this list of floats
			expected_length = 7
		else:
			raise ValueError(
				(
					'The line_type <{}> provided does not match the available options of:\n'
					'\t - {}\n'
					'\r - {}\n'
					'Check the code!'
				).format(line_type, self.current, self.impedance)
			)

		return cols, expected_length


class Loads:
	bus = 'NUMBER'
	load = 'MVAACT'
	identifier = 'ID'
	status='STATUS'
	zone='ZONE'

	default_id = '1'

	def __init__(self):
		"""
			Purely to avoid error codes
		"""
		pass


class Machines:
	bus = 'NUMBER'
	identifier = 'ID'
	rpos = 'RPOS'
	rneg = 'RNEG'
	rzero = 'RZERO'
	xsynch = 'XSYNCH'
	xtrans = 'XTRANS'
	xsubtr = 'XSUBTR'
	xneg = 'XNEG'
	xzero = 'XZERO'
	zsource = 'ZSORCE'
	rsource = 'R Source'
	xsource = 'X Source'

	t1d0 = "T'd0"
	t11d0 = "T''d0"
	t1q0 = "T'q0"
	t11q0 = "T''q0"

	xd = 'Xd'
	xq = 'Xq'
	x1d = "X'd"
	x1q = "X'q"
	x11 = "X''"

	tx_r = 'TX_R'
	tx_x = 'TX_X'

	# Minimum expected realistic RPOS value
	min_r_pos = 0.0
	# Assumed X/R value when they are missing
	assumed_x_r = 40.0

	bkdy_col_order = [bus, identifier, t1d0, t11d0, t1q0, t11q0, xd, xq, x1d, x1q, x11]

	# Defines the option for psspy.cong with regards to treatment of conventional machines and induction machines
	# 0 = Uses Zsorce for conventional machines
	# 1 = Uses X'' for conventional machines
	# 2 = Uses X' for conventional machines
	# 3 = Uses X for conventional machines
	bkdy_machine_type = 0

	def __init__(self):
		pass


class Plant:
	bus = 'NUMBER'
	status = 'STATUS'

	def __init__(self):
		"""
			Purely to avoid error messages
		"""
		pass


class Busbars:
	bus = 'NUMBER'
	state = 'TYPE'
	nominal = 'BASE'
	voltage = 'PU'
	bus_name = 'EXNAME'
	zone='ZONE'

	# Busbar type code lookup
	generator_bus_type_code = 2

	def __init__(self):
		pass


class Logging:
	"""
		Log file names to use
	"""
	logger_name = 'JK7938'
	debug = 'DEBUG'
	progress = 'INFO'
	error = 'ERROR'
	extension = '.log'

	def __init__(self):
		"""
			Just included to avoid Pycharm error message
		"""
		pass


class Excel:
	""" Constants associated with inputs from excel """
	circuit = 'Circuits'
	tx2 = '2 Winding'
	tx3 = '3 Winding'
	busbars = 'Busbars'
	fixed_shunts = 'Fixed Shunts'
	switched_shunts = 'Switched Shunts'
	machine_data = 'Machines'

	def __init__(self):
		pass


class G74:
	# Assumed X/R ratio of equivalent motor connected at 33kV
	x_r_33 = 2.76
	x_r_11 = 2.76
	# MVA contribution of equivalent motor per MVA of connected load (some ratio of these may be needed
	# based on whether load is assumed to be LV or HV connected.
	# NOTE - These values are not used and instead the values determined by SHETL are used
	label_mva = 'Machine Base'
	mva_lv = 1.0
	mva_hv = 2.6

	# 11 and 33kV parameters as per SHETL documentation
	# TODO: Validate SHETL parameters and document in report
	mva_33 = 1.16
	mva_11 = 1.16

	# Minimum MVA value for load to be considered
	min_load_mva = 0.15

	# Labels for DataFrame
	label_voltage = 'Load Voltage'
	hv = 'hv'
	lv = 'lv'

	machine_id = 'LD'

	# Time constants
	t11 = 0.04

	# Calculation of R and X'' for equivalent machine connected at 33kV and assumes
	# Z=1.0 which is then multiplied by the MVA rating of the machine
	rpos = (1.0/(1.0+x_r_33**2))**0.5
	x11 = (1.0-rpos**2)**0.5
	rzero = 10000.0
	xzero = 10000.0

	# Transformer impedance between 33kV and 11kV representation
	tx_r = 0.04
	tx_x = 0.6

	# Convert parameters to dictionary for easy updating in PSSe
	parameters_33 = {
			Machines.rpos: rpos,
			Machines.tx_r: 0.0,
			Machines.tx_x: 0.0,
			Machines.xsubtr: x11,
			Machines.xtrans: x11,
			Machines.xsynch: x11,
			Machines.rneg: rpos,
			Machines.xneg: x11,
			Machines.rzero: rzero,
			Machines.xzero: xzero,
			Machines.xsource: x11,
			Machines.rsource: rpos
		}

	# Parameters for 11kV connected loads that take into consideration the transformer between the
	# 33kV and 11kV busbars
	parameters_11 = {
		Machines.rpos: rpos-tx_r,
		Machines.tx_r: tx_r,
		Machines.tx_x: tx_x,
		Machines.xsubtr: x11-tx_x,
		Machines.xtrans: x11-tx_x,
		Machines.xsynch: x11-tx_x,
		Machines.rneg: rpos-tx_r,
		Machines.xneg: x11-tx_x,
		Machines.rzero: rzero,
		Machines.xzero: xzero,
		Machines.xsource: x11-tx_x,
		Machines.rsource: rpos-tx_r
	}

	# TODO: Calculate parameters for 33/11kV transformers and sensitivity study to determine the impact of these values
	# Transformer data in per unit on 100MVA base values
	# No longer accounting for 33/11kV transformers on a case by case basis but applying
	# SHETL parameters detailed above
	# #tx_r = 0.07142
	# #tx_x = 1.0

	# This is the minimum fault time that must be considered for the faults to determine Ik'' and Ip
	min_fault_time = 0.0
	# This is the time considered for returning the peak fault current
	peak_fault_time = 0.01

	def __init__(self):
		"""
			Purely to avoid error message
		"""
		pass


class SHEPD:
	"""
		Constants specific to the WPD study
	"""
	# voltage_limits are declared as a dictionary in the format
	# {(lower_voltage,upper_voltage):(pu_limit_lower,pu_limit_upper)} where
	# <lower_voltage> and <upper_voltage> represent the extremes over which
	# the voltage <pu_limit_lower> and <pu_limit_upper> applies
	# These are based on the post-contingency steady state limits provided in
	# EirGrid, "Transmission System Security and Planning Standards"
	# 380 used since that is base voltage in PSSE
	steady_state_limits = {
		(109, 111.0): (99.0 / 110.0, 120.0 / 110.0),
		(219.0, 221.0): (200.0 / 220.0, 240.0 / 220.0),
		(250.0, 276.0): (250.0 / 275.0, 303.0 / 275.0),
		(379.0, 401.0): (360.0 / 380.0, 410.0 / 380.0)
	}

	reactor_step_change_limit = 0.03
	cont_step_change_limit = 0.1

	# Unit used for fault times
	time_units = 'seconds'

	# The following headers will be used for the fault current output spreadsheet
	output_headers = ('Time after fault:', 'Value:')

	# SHEPD has a custom PSSE path installation which is defined here:
	psse_path = r'C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Siemens PTI\PSSE 33'

	# This is a threshold value, circuits with ratings less than this are reported and ignored
	rating_threshold = 0

	# Default time constant values to assume
	t1d0 = 0.12
	t11d0 = 0.04
	t1q0 = t1d0
	t11q0 = t11d0

	# Names and results associated with each type of result
	cb_make = 'Make'
	cb_break = 'Break'
	cb_steady = 'Steady'
	results_per_fault = dict()
	results_per_fault[cb_make] = [
		BkdyFileOutput.ik11,
		BkdyFileOutput.ip,
		BkdyFileOutput.x,
		BkdyFileOutput.r
	]
	results_per_fault[cb_break] = [
		BkdyFileOutput.ibsym,
		BkdyFileOutput.ibasym
	]

	cols_for_min_fault_time = [
		BkdyFileOutput.ik11,
		BkdyFileOutput.ibsym,
		BkdyFileOutput.ibasym,
		BkdyFileOutput.idc,
		BkdyFileOutput.x,
		BkdyFileOutput.r
	]

	cols_for_peak_fault_time = [
		BkdyFileOutput.ip,
		BkdyFileOutput.ibsym,
		BkdyFileOutput.ibasym,
		BkdyFileOutput.idc
	]

	cols_for_other_fault_time = [
		BkdyFileOutput.ibsym,
		BkdyFileOutput.ibasym,
		BkdyFileOutput.idc
	]

	# List controls the order of the output columns for the LTDS export
	output_column_order = [
		BkdyFileOutput.ik11,
		BkdyFileOutput.ip,
		BkdyFileOutput.ibsym,
		BkdyFileOutput.r,
		BkdyFileOutput.x
	]

	def __init__(self):
		""" Purely added to avoid error message"""
		pass


class LTDS:
	bus_name = 'Name'
	nominal = 'Voltage (kV)'

	def __init__(self):
		""" Purely added to avoid error message"""
		pass


class PowerFactory:
	"""
		Constants used in this script
	"""
	# Constants relating to the paths
	pf_year = 2019
	year_max_tested = 2019
	pf_service_pack = ''
	dig_path = str()
	dig_python_path = str()
	# Populated with available installed PowerFactory versions on initialisation
	available_power_factory_versions = list()
	target_power_factory = 'PowerFactory 2019'

	# The following list details python versions which are non compatible
	non_compatible_python_versions = ['3.5']

	# Default PowerFactory installation directories
	default_install_directory = r'C:\Program Files\DIgSILENT'
	power_factory_search = 'PowerFactory 20*'

	# String used to define the tuning frequency of the filter
	hz='Hz'
	sht_Filters = 'Filters'
	sht_Terminals = 'Terminals'
	sht_Scenarios = 'Base_Scenarios'
	sht_Contingencies = 'Contingencies'
	sht_Study = 'Study_Settings'
	sht_LF = 'Loadflow_Settings'
	sht_Freq = 'Frequency_Sweep'
	sht_HLF = 'Harmonic_Loadflow'
	# Different filter types available in PowerFactory 2016
	Filter_type = {'C-Type':4,
				   'Single':0,
				   'High Pass':3}
	pf_substation = 'ElmSubstat'
	pf_line = 'ElmLne'
	pf_branch = 'ElmBranch'
	pf_terminal =  'ElmTerm'
	pf_coupler = 'ElmCoup'
	pf_mutual = 'ElmMut'
	pf_fault_event = 'IntEvt'
	pf_switch_event = 'EvtSwitch'
	pf_outage_event = 'EvtOutage'
	pf_case = 'IntCase'
	pf_scenario = 'IntScenario'
	pf_filter = 'ElmShnt'
	pf_cubicle = 'StaCubic'
	pf_term_voltage = 'uknom'
	pf_shn_term = 'bus1'
	pf_shn_voltage = 'ushnm'
	pf_shn_type = 'shtype'
	pf_shn_q = 'qtotn'
	pf_shn_inputmode = 'mode_inp'
	pf_shn_selectedinput = 'DES'
	pf_shn_freq = 'fres'
	pf_shn_tuning = 'nres'
	pf_shn_qfactor = 'greaf0'
	pf_shn_qfactor_nom = 'grea'
	pf_shn_rp = 'rpara'
	# constants for variations
	pf_scheme = 'IntScheme'
	pf_stage = 'IntSstage'
	pf_results = 'ElmRes'
	pf_network_elements = 'ElmNet'
	pf_project = 'IntPrj'
	# Command for carrying out contingency analysis and applying each outage
	pf_cont_analysis = 'ComSimoutage'
	pf_outage = 'ComOutage'

	# General Types
	pf_folder_type = 'IntFolder'
	pf_fault_cases_folder = 'IntFltcases'
	pf_netdata_folder_type = 'netdat'
	pf_faults_folder_type = 'fault'
	pf_sc_folder_type = 'study'
	pf_os_folder_type = 'scen'

	# Default results file name
	default_fs_extension = '_FS'

	pf_r1 = 'm:R'
	pf_x1 = 'm:X'
	pf_z1 = 'm:Z'
	pf_r12 = 'c:R_12'
	pf_x12 = 'c:X_12'
	pf_z12 = 'c:Z_12'
	pf_nom_voltage = 'e:uknom'
	pf_freq = 'b:fnow in Hz'
	pf_harm = 'b:ifnow'

	ldf_command = 'ComLdf'
	hldf_command = 'ComHldf'
	fs_command = 'ComFsweep'
	autotasks_command = 'ComTasks'

	# Folder names for temporary folders
	temp_sc_folder = 'temp_sc'
	temp_os_folder = 'temp_os'
	temp_faults_folder = 'temp_faults'
	temp_mutual_folder = 'mutual_elements'

	# Constants associated with the handling of PowerFactory initialisation and
	# potential intermittent errors
	# Number of attempts to obtain a license
	license_activation_attempts = 5
	# Number of seconds to wait between license attempts
	license_activation_delay = 5.0
	# Error codes which could be intermittent and therefore the script should try again
	# Description in PowerFactory help file:  ErrorCodeReference_en.pdf
	license_activation_error_codes = (3000, 3002, 3005, 3011, 3012, 4000, 4002, 5000)

	# Each results variable has a default type and need to assign the defaults to the newly created results
	# variables
	def_results_hlf = 5		# Harmonic load flow
	def_results_fs = 9		# Frequency sweep
	def_results_cont = 13#

	# User default settings
	user_default_settings = 'Set\Def\Settings.SetUser'

	# Number of seconds to allow when waiting for parallel processor response
	parallel_time_out = 100

	# This is a maximum impedance value, above this and it is assumed to be open circuit and will be ignored
	max_impedance = 1E6

	class ComRes:
		# Power Factory class name
		pf_comres = 'ComRes'
		#Com Res setting constants

		# File export type:
		#	0 = Output window
		#	1 = Windows clipboard
		#	2 = Measurement file (ElmFile)
		#	3 = Comtrade
		#	4 = Testfile
		#	5 = PSSPLT Version 2.0
		#	6 = Commas Separated Values (*.csv)
		export_type = 'iopt_exp'
		# Name of file to export to (if appropriate)
		file = 'f_name'
		# Type of separators to use (0 = Custom, 1 = system defaults)
		separators = 'iopt_sep'
		# Export object headers only (0 = all data, 1 = headers only)
		object_head_only = 'iopt_honly'
		# Variables to extract (0 = all, 1 = custom list)
		variables_all = 'iopt_csel'
		# Name of result file from PF to export
		result = 'pResult'
		# Details to export from element:
		# 	0 = None,
		# 	1 = Name,
		# 	2 = Short path and name,
		# 	3 = Path and name,
		# 	4 = Foreign key
		element = 'iopt_locn'
		# Details to export from variable:
		#	0 = None,
		#	1 = Parameters name,
		#	3 = Short description,
		#	4 = Full description
		variable = 'ciopt_head'
		# Custom of full dataset (0 = full, 1 = custom)
		user_interval = 'iopt_tsel'
		# Export values (0 = values, 1 = variable descriptors only)
		export_values = 'iopt_vars'
		# Shift time of results (0 = none, 1 = shift)
		shift_time = 'iopt_rscl'
		# Filter time of results (0 = None, 1 = filter)
		filtered_time = 'filtered'

	def __init__(self):
		"""
			Initialises the relevant python paths depending on the version that has been loaded
		"""
		# Get reference to logger
		self.logger = logger

		# Find all PowerFactory versions installed in this location
		power_factory_paths = glob.glob(os.path.join(self.default_install_directory, self.power_factory_search))
		self.available_power_factory_versions = [os.path.basename(x) for x in power_factory_paths]
		self.available_power_factory_versions.sort()

	def select_power_factory_version(self, pf_version=None, mock_python_version=str()):
		"""
			Function allows the user to select a specific PowerFactory version, if none is selected then
			the most recent version of PowerFactory is used
		:param str pf_version: (optional) - If None then the most recent PowerFactory version is used
		:param str mock_python_version:  For TESTING only, gets replaced with a different version to check correct
										errors thrown if incorrect version provided
		:return str self.target_power_factory:  Returns the PowerFactory version in string
		"""

		# If no pf_version is provided then the default version defined is used if it exists in the available versions
		# otherwise the latest version
		if pf_version is None:
			if self.target_power_factory not in self.available_power_factory_versions:
				# Rather than assuming a particular version just default to the latest version
				self.target_power_factory = self.available_power_factory_versions[-1]
		elif pf_version in self.available_power_factory_versions:
			self.target_power_factory = pf_version
		else:
			self.logger.critical(
				(
					'The PowerFactory version {} has been selected but does not exist in the installed versions:\n\t{}'
				).format(pf_version, '\n\t'.join([x for x in self.available_power_factory_versions]))
			)
			raise EnvironmentError('Invalid PowerFactory version')

		self.logger.debug('PowerFactory version <{}> will be used'.format(self.target_power_factory))

		# Find year from selected PowerFactory version
		# pf_version is assumed to take the format PowerFactory #### and therefore the #### can be extracted
		year = [int(s) for s in self.target_power_factory.split() if s.isdigit()][0]

		# Confirm the year is > 2017 and < 2020 otherwise warn that hasn't been fully tested
		if int(year) < 2018:
			self.logger.warning(
				(
					'You are using PowerFactory version {}.\n'
					'In the 2018 version there were some API changes which have been considered in this script.  The '
					'previous versions may still work but they have not been considered as part of the development '
					'testing and so you are advised to carefully check your results.'
				).format(year)
			)
		elif int(year) > self.year_max_tested:
			self.logger.warning(
				(
					'You are using PowerFactory version {}.\n'
					'This script has only been tested up to year {} and therefore changes in the PowerFactory API may '
					'impact on the running and results you produce.  You are advised to check the results carefully or '
					'consider updating the developments testing for this version.  For further details contact:\n{}'
				).format(year, self.year_max_tested, Author.contact_summary)
			)

		# Find the installation directory for the PowerFactory paths
		self.dig_path = os.path.join(self.default_install_directory, self.target_power_factory)

		# Now checks for Python versions within this PowerFactory
		if mock_python_version:
			# Running in a test mode to check an error is created
			self.logger.warning('TESTING - Testing with a mock python version to raise exception, if not expected then there is a '
								'script error! - TESTING')
			current_python_version = mock_python_version
		else:
			# Formulate string for python version
			current_python_version = '{}.{}'.format(sys.version_info.major, sys.version_info.minor)


		# Get list of supported python versions
		list_of_available_versions = [os.path.basename(x) for x in glob.glob(os.path.join(self.dig_path, 'Python', '*'))]
		if current_python_version in self.non_compatible_python_versions:
			self.logger.critical(
				(
					'You are using Python version {}, this script is not compatible with that version or the following '
					'versions: \n\t Python {}\n  Additionally, the PowerFactory version you have selected ({}) is only compatible '
					'with the following Python versions: \n\t Python {}'
				).format(
					current_python_version, '\n\t Python '.join(self.non_compatible_python_versions),
					self.target_power_factory, '\n\t Python '.join(list_of_available_versions)
				)
			)
			raise EnvironmentError('Non Compatible Python Version')

		# Define the Python Path for PowerFactory
		self.dig_python_path = os.path.join(self.dig_path, 'Python', current_python_version)
		if not os.path.isdir(self.dig_python_path):

			self.logger.critical(
				(
					'You are running python version: {} but only the following versions are supported by this version of'
					'PowerFactory ({}):\n\t Python {}'
				).format(current_python_version, self.target_power_factory, '\n\t Python '.join(list_of_available_versions))
			)
			raise EnvironmentError('Incompatible Python version')

		# Returns the PF version that is being initialised
		return self.target_power_factory


class LociInputs:
	# Default values to use for impedance loci processing
	# Defaults to using +/- half of nominal frequency
	def_polygon_range = General.nominal_frequency / 2.0
	# Defaults to not excluding any data points
	def_impedance_exclude = 0.0

	# Maximum harmonic order (no real impact just avoids excessive loops)
	max_harm = 100

	# Strings defined in inputs
	unlimited_inputs = 'Unlimited'
	custom_inputs = 'Custom'
	min_frequency = 'Minimum Frequency (Hz)'
	max_frequency = 'Maximum Frequency (Hz)'
	percentage_to_exclude = 'Percentage to Exclude (%)'
	max_vertices = 'Maximum No. Vertices'

	# Default values to use for if no loci vertice restrictions are in place
	unlimited_identifier = 10000
	def_max_vertices = unlimited_identifier

	# Minimum allowable number of vertices, less than this and no calculation is really possible
	min_vertices = 4
	max_num_vertices = 100

	# This is the percentage of the impedance that the vertice will be moved by, the smaller this is
	# the longer results processing will take but the less likely that the loci will be increased in
	# size excessively
	vertice_step_size = 0.1