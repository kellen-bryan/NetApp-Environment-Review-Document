# FILE: 	NERD.py
#
# PROGRAM:	NetApp Environment Review Document (NERD)
# 
# AUTHOR: Kellen Bryan
#  
# SUMMARY: NERD class for NERD_modeler.py
#
# Copyright (c) 2017 Network Appliance, Inc.
# All rights reserved.

########## MODULE IMPORT ############################################## 

#alphabet
import fileinput, sys, re, codecs
import argparse 
import os.path
import requests
import statistics
from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill, Font, Fill, Alignment, Border, Side
from openpyxl.worksheet.properties import WorksheetProperties, PageSetupProperties
from openpyxl.utils import column_index_from_string, get_column_letter
from dateutil.relativedelta import relativedelta
from datetime import datetime,tzinfo,date
from operator import itemgetter

class NERD():

	serial_numbers = []

	def __init__(self, serial_numbers_list):
		"""initialize Environment by adding serial #'s' to list. Serial #'s are accepted as argument or file"""
		for num in serial_numbers_list:
			self.serial_numbers.append(num)

	def _aggr_capacity(self, asup_url_output):
		"""Returns aggregate capacity"""
		line = asup_url_output
		match = re.findall("aggr_allocated_kb>(.*?)</aggr_allocated_kb>", line, re.DOTALL)
		match_name = re.findall("aggr_name>(.*?)</aggr_name>", line, re.DOTALL)
		match_return 	= {} #dictionary -> Key: aggr name; Value: aggr capacity

		if match:

			length = len(match)
			
			for i in range(0, length):
				capacity_TB = (float(match[i]) / 1000000000)
				capacity_TB_rounded = round(capacity_TB, 2)
				match_return[str(match_name[i])] = capacity_TB_rounded
			
			return match_return

	def _aggr_name(self, asup_url_output):
		"""Returns aggregate names"""
		line = asup_url_output
		match = re.findall("(?<=Aggregate )(.*?)(?= \()", line)

		if match:
			match_list = []
			for name in match:
				if name not in match_list:
					match_list.append(str(name))
				elif name in match_list:
					break

			return match_list

	def _aggr_util(self, asup_url_output):
		"""Returns percent of aggr utilized"""
		line = asup_url_output
		match = re.findall("aggr_used_pct>(.*?)</aggr_used_pct>", line, re.DOTALL)
		match_name = re.findall("aggr_name>(.*?)</aggr_name>", line, re.DOTALL)
		match_return 	= {} #dictionary -> Key: aggr name; Value: aggr capacity

		if match:
			length = len(match)
			
			for i in range(0, length):
				match_return[str(match_name[i])] = int(match[i])
			
			return match_return

	def _asup_id(self, asup_url_output):
		"""Returns asup id for rest API use"""
		line = asup_url_output
		match = re.search("asup_id>(.*?)</asup_id", line, re.DOTALL)
		if match:
			return str(match.group(1))
	
	def _asup_received_date(self, asup_url_output):
		"""Return date of last recieved ASUP data"""
		line = asup_url_output
		match = re.search("asup_received_date>(.*?)</asup_received_date>", line, re.DOTALL)
		if match:
			return str(match.group(1))

	def _biz_key(self, asup_url_output):
		"""Returns biz_key for rest api use"""
		line = asup_url_output
		match = re.search("biz_key>(.*?)</biz_key>", line, re.DOTALL)
		if match:
			return str(match.group(1))

	def _cluster_name(self, asup_url_output):
		"""Returns client-defined cluster name"""
		line = asup_url_output
		match = re.search("cluster_name>(.*?)<", line)
		if match:
			return str(match.group(1))

	def _disk_count(self, asup_url_output):

		line = asup_url_output
		match = re.search("<data>(.*?)</data>", line, re.DOTALL)
		match_name = re.findall("(?<=Aggregate )(.*?)(?= \()", line)

		if match:

			match_buffer 	= []
			match_buffer 	= re.findall("(\S+)", str(match.group(1)))
			match_return 	= {} #dictionary -> Key: aggr name; Value: plex names

			disk_counter	= 0
			raid_group_flag = 0
			name_value 		= -1
			name_list 		= []

 			#add unique aggr name keys and create name list
			for agg_name in match_name:
				if agg_name not in name_list:
					name_list.append(str(agg_name))
					match_return[str(agg_name)] = 0
				elif agg_name in match_buffer:
					break
			
			#find disk count for each aggr
			for word in match_buffer:
				if "parity" in word and raid_group_flag == 1 and name_value < len(name_list):
					match_return[name_list[name_value]] = str(disk_counter) + " disks"
					raid_group_flag = 2
					disk_counter = 0

				elif "data" in  word and raid_group_flag != 2:
					raid_group_flag = 1
					disk_counter += 1

				elif "Aggregate" in word:
					if raid_group_flag == 1:
						raid_group_flag = 0
						match_return[name_list[name_value]] = str(disk_counter) + " disks"
						disk_counter = 0
						name_value += 1
						
					else:
						name_value += 1
						raid_group_flag = 0
					

				elif "spare" in word and raid_group_flag != 2:
					match_return[name_list[name_value]] = str(disk_counter) + " disks"
					break
			
			return match_return

	def _host_name(self, asup_url_output):
		"""Returns product host name"""
		line = asup_url_output
		match = re.search("hostname>(.*?)<", line)
		if match:
			return str(match.group(1))

	def _location(self, asup_url_output):
		"""Returns customer location. Location defines different pages"""
		line = asup_url_output
		match = re.search("site_name>(.*?)<", line)
		if match:
			return str(match.group(1))

	def _performance_iops(self, asup_url_output):
		"""Returns iops and cpu busy % (std_dev + avg) for previous week"""
		line = asup_url_output
		fcp_ops_match = re.findall("fcp_ops.*?<counterValue>(.*?)</counterValue>", line, re.DOTALL)
		iscsi_ops_match = re.findall("iscsi_ops.*?<counterValue>(.*?)</counterValue>", line, re.DOTALL)
		cifs_ops_match = re.findall("cifs_ops.*?<counterValue>(.*?)</counterValue>", line, re.DOTALL)
		nfs_ops_match = re.findall("nfs_ops.*?<counterValue>(.*?)</counterValue>", line, re.DOTALL)
		cpu_busy_match = re.findall("cpu_busy.*?<counterValue>(.*?)</counterValue>", line, re.DOTALL)

		if fcp_ops_match or iscsi_ops_match or cifs_ops_match or nfs_ops_match or cpu_busy_match:

			fcp_ops 	= []
			iscsi_ops 	= []
			cifs_ops 	= []
			nfs_ops 	= []
			cpu_busy 	= []

			length = len(fcp_ops_match)
			
			for i in range(0, length):
				fcp_ops.append(float(fcp_ops_match[i]))
				iscsi_ops.append(float(iscsi_ops_match[i]))		
				cifs_ops.append(float(cifs_ops_match[i]))		
				nfs_ops.append(float(nfs_ops_match[i]))		
				cpu_busy.append(float(cpu_busy_match[i]))		

			#calculate averages and std dev over last week
			fcp_avg 		= sum(fcp_ops)/length
			fcp_std_dev 	= statistics.stdev(fcp_ops)

			iscsi_avg 		= sum(iscsi_ops)/length
			iscsi_std_dev 	= statistics.stdev(iscsi_ops)

			cifs_avg 		= sum(cifs_ops)/length
			cifs_std_dev 	= statistics.stdev(cifs_ops)

			nfs_avg		 	= sum(nfs_ops)/length
			nfs_std_dev 	= statistics.stdev(nfs_ops)

			cpu_avg 		= sum(cpu_busy)/length
			cpu_std_dev 	= statistics.stdev(cpu_busy)

			return_list = []
			return_list.append(round( (cpu_avg+cpu_std_dev), 2))
			return_list.append(round( (cifs_avg+cifs_std_dev), 2))
			return_list.append(round( (fcp_avg+fcp_std_dev), 2))
			return_list.append(round( (iscsi_avg+iscsi_std_dev), 2))
			return_list.append(round( (nfs_avg+nfs_std_dev), 2))
			
			return return_list

		no_match_list = ['No Data Available', 'No Data Available', 'No Data Available', 'No Data Available', 'No Data Available']
		return no_match_list

	def _raid_group_count(self, asup_url_output):
		"""Returns RAID group set-up"""
		line = asup_url_output
		match_name = re.findall("(?<=Aggregate )(.*?)(?= \()", line)
		match_raid = re.findall("(?<=RAID group)(.*?)(?=\()", line)

		if match_name and match_raid:
			match_buffer = {} #dictionary -> Key: aggr name; Value: raid group names
			for agg_name in match_name:
				if agg_name not in match_buffer: #add unique aggr name keys
					match_buffer[str(agg_name)] = [] 
					for raid in match_raid:
						if agg_name in raid and raid not in match_buffer[str(agg_name)]: #add plex names to corresponding aggr names
							match_buffer[str(agg_name)].append(str(raid))
				elif agg_name in match_buffer:
					break
			
			return match_buffer
	
	def _raid_type(self, asup_url_output):
		"""Returns percent of aggr utilized"""
		line = asup_url_output
		match = re.findall("aggr_raid_type>(.*?)</aggr_raid_type>", line, re.DOTALL)
		match_name = re.findall("aggr_name>(.*?)</aggr_name>", line, re.DOTALL)
		match_return 	= {} #dictionary -> Key: aggr name; Value: aggr capacity

		if match:
			length = len(match)
			
			for i in range(0, length):
				match_return[str(match_name[i])] = str(match[i])
			
			return match_return

	def _disk_type_count(self, asup_url_output):
		"""Returns number of each disk type"""
		line = asup_url_output
		match = re.search("<data>(.*?)</data>", line, re.DOTALL)
		match_name = re.findall("(?<=Aggregate )(.*?)(?= \()", line)

		if match:

			match_buffer 	= []
			match_buffer 	= re.findall("(\S+)", str(match.group(1)))
			match_return 	= {} #dictionary -> Key: aggr name; Value: RAID Group count

			name_value 		= 0
			name_list 		= []
			ssd_counter 	= 0
			sas_counter 	= 0
			new_aggr_flag 	= 0
			start_flag 		= 0
			i 				= 0

 			#add unique aggr name keys and create name list
			for agg_name in match_name:
				if agg_name not in name_list:
					name_list.append(str(agg_name))
					match_return[str(agg_name)] = 0
				elif agg_name in match_buffer:
					break

			#Count number of each type of RAID layout
			for word in match_buffer:

				i+=1

				if "Aggregate" in word or "spare" in word:
					if start_flag != 0:
						new_aggr_flag = 0
						match_return[name_list[name_value]] = "({} SAS; {} SSD)".format(sas_counter, ssd_counter)
						name_value += 1
						if "spare" in word:
							break
					else:
						start_flag = 1

				elif "Type" in word:
					if new_aggr_flag == 0:
						sas_counter = 0
						ssd_counter = 0
						new_aggr_flag = 1

					raid_type = match_buffer[i + 32]
					if "SSD" in raid_type:
						ssd_counter += 1
					elif "SAS" in raid_type:
						sas_counter += 1

			return match_return

	def _system_model(self, asup_url_output):
		"""Returns system model/controller"""
		line = asup_url_output
		match = re.search("sys_model>(.*?)<", line)
		if match:
			return str(match.group(1))

	def _serial_number(self, asup_url_output):
		"""Returns product serial number"""
		line = asup_url_output
		match = re.search("sys_serial_no>(.*?)<", line)
		if match:
			return str(match.group(1))

	def _system_version(self, asup_url_output):
		"""Returns OS version"""
		line = asup_url_output
		match = re.search("sys_version>(.*?)<", line)
		if match:
			return str(match.group(1))	

	def _warranty_status(self, asup_url_output):
		print "Warranty Status"
		"""Returns warranty expiration date"""

	
			
	

	

	

	

	

#dont know
#http://restprd.corp.netapp.com/asup-rest-interface/ASUP_DATA/client_id/test/biz_key/C%7C93E0D750-5BDF-11E4-9F60-123478563412%7C8499809755%7C721545000241/object_list

#full list
#http://restprd.corp.netapp.com/asup-rest-interface/ASUP_DATA/client_id/test/biz_key/C%7C93E0D750-5BDF-11E4-9F60-123478563412%7C8499809755%7C721545000241/l