
import os
import constants as constants
import glob
import pandas as pd
import numpy as np
import shapely.geometry
import collections
import math
import shutil
import time
import xlsxwriter
import xlsxwriter.utility
import matplotlib.pyplot
import common_functions as common
import random


def cross_sign(x1, y1, x2, y2):
	"""
		True if cross is positive, false if negative or zero
	:param float x1:
	:param float y1:
	:param float x2:
	:param float y2:
	:return bool status:
	"""
	# True if cross is positive
	# False if negative or zero
	status = x1 * y2 > x2 * y1
	return status

def angle(x1, y1, x2, y2):
	"""
		Use dotproduct to find the angle between vectors.

		Angle returned in degrees between 0, 180 degrees
	:param float x1:
	:param float y1:
	:param float x2:
	:param float y2:
	:return float angle:
	"""
	# Calculate numerator
	numerator = (x1 * x2 + y1 * y2)
	# Calculate denominator
	denominator = math.sqrt((x1 ** 2 + y1 ** 2) * (x2 ** 2 + y2 ** 2))
	an = math.degrees(math.acos(numerator / denominator))
	return an


def new_coordinates(x_source, y_source, x_target, y_target):
	"""
		Calculate a new coordinate for the loci point extending in the
		direction of existing line to straighten out the points for the
		parameter with the closest angle
	:param float x_source:  x co-ordinate of start of line
	:param float y_source:  y co-ordinate of start of line
	:param float x_target:  x co-ordinate of point to be adjusted
	:param float y_target:  y co-ordinate of point to be adjusted
	:return (float, float), (x_new, y_new):  Returns new coordinates for x and y values
	"""

	# Determine length of source line
	source_length = (((y_target-y_source)**2+(x_target-x_source)**2)**0.5)*constants.LociInputs.vertice_step_size

	# Calculate the deltas between the source point and the target point
	y_change = y_target-y_source
	x_change = x_target-x_source

	# To avoid a div/0 error ensure a minimum value exists since result will be 90 degrees
	if x_change == 0.0:
		x_change = 0.0001

	# Angle of line and determine direction
	beta = math.atan(y_change/x_change)
	if beta < 0:
		y_multiplier = -1
	else:
		y_multiplier = 1

	# Calculate delta changes in x and y directions
	dx = abs(source_length*math.cos(beta))
	dy = abs(source_length*math.sin(beta))*y_multiplier

	# To account for quadrant calculate in both directions
	x_new1 = x_target + dx
	y_new1 = y_target + dy

	x_new2 = x_target - dx
	y_new2 = y_target - dy

	# Determine distance from new points to source
	length1 = (((y_new1-y_source)**2+(x_new1-x_source)**2)**0.5)*constants.LociInputs.vertice_step_size
	length2 = (((y_new2-y_source)**2+(x_new2-x_source)**2)**0.5)*constants.LociInputs.vertice_step_size
	lengths = (length1, length2)

	# Determine the longest length which means it is in the correct direction and then implement that
	idx_max = lengths.index(max(lengths))
	if idx_max == 0:
		x_new = x_new1
		y_new = y_new1
	else:
		x_new = x_new2
		y_new = y_new2

	return x_new, y_new



def find_convex_vertices(x_values, y_values, max_vertices, node='None', h='None'):
	"""
		Finds the ConvexHull that bounds around the provided x and y values and ensures that the maximum number
		of vertices does not exceed the provided value
	:param tuple x_values: X axis values to be considered
	:param tuple y_values: Y axis values to be considered
	:param int max_vertices: Maximum number of vertices to allow
	:param str node:  Name of node for logging purposes
	:param str h:  Harmonic number for logging purposes
	:return tuple corners: (x / y points for each corner
	"""
	c = constants.PowerFactory
	# Constants used for columns in DataFrame
	lbl_x = 'x'
	lbl_y = 'y'
	lbl_an = 'angle'

	# # Filter out values which are outside of allowed range
	# x_points = list()
	# y_points = list()
	# for x,y in zip(x_values, y_values):
	# 	if 0 < abs(x) < c.max_impedance and 0 < abs(y) < c.max_impedance:
	# 		x_points.append(x)
	# 		y_points.append(y)

	# Convert provided data points into a MultiPoint object for only those x and y values that are within the acceptable
	# range of being greater than 0 and less than the maximum allowed impedance
	all_points = shapely.geometry.MultiPoint(
		[
			shapely.geometry.Point((x,y)) for x,y in zip(x_values, y_values)
			if 0 < abs(x) < c.max_impedance and 0 < abs(y) < c.max_impedance
		 ]
	)

	# Check that MultiPoint array is not empty, if it is then return none
	if all_points.is_empty:
		return list(), list()

	# Determine convex hull and number of vertices (subtracting 1 to account for returning to the start)
	convex_hull = all_points.convex_hull

	if type(convex_hull) is shapely.geometry.Point:
		# Convex hull is a single point and therefore just need to return the points
		x_corner, y_corner = convex_hull.coords.xy
		return x_corner, y_corner
	elif type(convex_hull) is shapely.geometry.LineString:
		# Convex hull is a single line and therefore just need to return the points
		x_corner, y_corner = convex_hull.coords.xy
		return x_corner, y_corner


	x_corner, y_corner = convex_hull.exterior.xy
	num_vertices = len(x_corner)-1

	# Determine whether any limit on the number of vertices, if none then return corners
	if max_vertices >= constants.LociInputs.unlimited_identifier:
		return x_corner, y_corner

	# Direction flag alternates for each loop to ensure that overall polygon is expanded rather than just 1 corner
	direction = False

	# plt.scatter(x_values, y_values)
	# plt.plot(x_corner, y_corner, '-.', color='r')
	# plt.pause(0.01)

	# Loop to ensure the maximum number of vertices are not exceeded
	counter = 0
	counter_limit = 1E6
	# num_vertices must be greater than 4 otherwise dealing with an envelope but the minimum number of allowed vertices
	# will always be greater than 4 so this cycle doesn't start
	while num_vertices>max_vertices and counter < counter_limit:
		# Counter just to ensure don't get stuck in infinite loop
		counter += 1
		# Direction toggles each time to ensure overall polygon expanded
		direction = not direction

		# Remove the last point of the corner which is the same as the starting point
		x = x_corner[:-1]
		y = y_corner[:-1]

		# Determine new expanded ConvexHull with vertices stored in DataFrame along with angle between vertices
		df = pd.DataFrame()
		for i in range(len(x)):
			# Obtain the reference points so vertice can be calculated
			point1 = (x[i], y[i])
			point_ref = (x[i-1], y[i-1])
			point2 = (x[i-2], y[i-2])

			# Calculate vertices
			x1, y1 = point1[0]-point_ref[0], point1[1]-point_ref[1]
			x2, y2 = point2[0]-point_ref[0], point2[1]-point_ref[1]

			# Add reference points to DataFrame
			df.loc[i, lbl_x] = point_ref[0]
			df.loc[i, lbl_y] = point_ref[1]

			# Calculate the angle between the vertices
			if cross_sign(x1,y1,x2,y2):
				# If clockwise then angle is negative for inner angle
				df.loc[i, lbl_an] = 180 - angle(x1,y1,x2,y2)
			else:
				df.loc[i, lbl_an] = angle(x1,y1,x2,y2)

		# Calculate the point either side of the one with the greatest angle depending on the direction for
		# this iteration
		# Point to be filtered out
		idx0 = df.loc[:, lbl_an].idxmax()
		if direction:
			# If in one direction then obtain value 2 steps round
			idx_target = idx0-1
			# Check that haven't breached all the index values and if so restart at other extreme
			if idx_target < min(df.index):
				idx_target = max(df.index)

			# Obtain source for line
			idx_source = idx_target-1
			# Confirm not outside of index limits
			if idx_source < min(df.index):
				idx_source = max(df.index)

		else:
			idx_target = idx0+1

			# Check that haven't breached all the index values and if so restart at other extreme
			if idx_target > max(df.index):
				idx_target = min(df.index)

			# Obtain source for line
			idx_source = idx_target+1
			# Confirm not outside of index limits
			if idx_source > max(df.index):
				idx_source = min(df.index)

		# Obtain the coordinate that should be moved and angle it should be moved at
		x_target, y_target = df.loc[idx_target, [lbl_x, lbl_y]]
		x_source, y_source = df.loc[idx_source, [lbl_x, lbl_y]]

		# Update the dataframe with the new x and y values
		df.loc[idx_target, [lbl_x, lbl_y]] = new_coordinates(
			x_source=x_source, y_source=y_source, x_target=x_target, y_target=y_target
		)
		# Obtain the new points that should be used for the convex hull
		new_points = shapely.geometry.MultiPoint(
			[shapely.geometry.Point(z) for z in df.loc[:, [lbl_x, lbl_y]].values]
		)

		# Calculate new ConvexHull corners, new vertices, etc.
		x_corner, y_corner = new_points.convex_hull.exterior.xy
		num_vertices = len(x_corner)-1
		# convex_hull = new_points.convex_hull

		# plt.plot(x_corner, y_corner, '-.', color='g')
		# plt.pause(0.001)

	# plt.plot(x_corner, y_corner, '-', color='b')
	# plt.show()
	# plt.clf()

	# if counter >= counter_limit:
	# 	constants.logger.error(
	# 		(
	# 			'In attempting to limit the vertices to {} ended up taking {} iterations and therefore abandoned.  '
	# 			'Results for node {} at harmonic number {} therefore has {} vertices'
	# 		).format(max_vertices, counter, node, h, num_vertices)
	# 	)

	# Return tuple of R and X values
	return x_corner, y_corner


if __name__ == '__main__':


    FILE_NAME_INPUT_1 = 'scenario_names_dig.xlsx'
    FILE_NAME_INPUT_1 = 'scenario_names_dig_BLE.xlsx'
    Project_name='Test1'
    Project_name = '20210225-OCW01-New_VImp_Main_BLE_test1'
    Project_name = 'New_OC_BLE_Filter1_test1'

    upper_limit = int(constants.PowerFactory.max_impedance - 1)
    number_points = 50
    PLOT_FIGURES = True

    x_points = (random.sample(range(upper_limit), number_points))
    y_points = (random.sample(range(upper_limit), number_points))

    corners = find_convex_vertices(
        x_values=x_points, y_values=y_points, max_vertices=constants.LociInputs.unlimited_identifier
    )

    # Confirm all points lie within the Polygon returned by the vertices
    polygon = shapely.geometry.polygon.Polygon(list(zip(*corners)))
    rand_point = random.randint(0, number_points - 1)
    point = shapely.geometry.Point(x_points[rand_point], y_points[rand_point])

    if PLOT_FIGURES:
        matplotlib.pyplot.plot(x_points, y_points, 'o')

        matplotlib.pyplot.plot(corners[0], corners[1], 'r--')
        matplotlib.pyplot.show()




    FILE_PTH_INPUT_1 = common.get_local_file_path_withfolder(file_name=FILE_NAME_INPUT_1, folder_name='dig_results')
    sc_dig_df = common.import_excel(FILE_PTH_INPUT_1)
    scenario_dig_list = list(sc_dig_df['scenario'])
    output_excel_dig_list = ['{}_{}'.format(x, 'points.xlsx') for x in scenario_dig_list]
    scenario_dig_list_V = ['{}_{}'.format(x, 'V') for x in scenario_dig_list]
    scenario_dig_list_I = ['{}_{}'.format(x, 'I') for x in scenario_dig_list]
    dig_scenarios=scenario_dig_list_V + scenario_dig_list_I

    pf_version = '2020'


