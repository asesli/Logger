import collections
from datetime import datetime, timedelta
import getpass
import itertools
import os
##PSUTIL ERRORS OUT ON WIN10/py38
import psutil
import win32com.client
import pythoncom
import re
from shutil import copyfile
from socket import gethostname
#import threading
#import thread
import win32api
import win32con
import win32gui
import win32process
import win32security
import wmi
import ftrack_api
#import anydbm 
#import dbhash
import gc


#os.environ['REQUESTS_CA_BUNDLE'] = 'C:/Python27/Lib/site-packages/certifi/cacert.pem' #copy this file to the dir below
os.environ['REQUESTS_CA_BUNDLE'] = 'L:/HAL/LIVEAPPS/utils/WorkTracker/certifi/cacert.pem'

class process_log():
	def __init__(self):
		self.user = getpass.getuser()
		self.log = "{0}log_{1}_{2}.txt".format("L:/HAL/LIVEAPPS/utils/WorkTracker/_slog/", self.user, gethostname())

		self.verbose = False

		if self.user.lower() not in ['render', 'system', 'adminuse']:
			self.processed = self.read_slog()
			
			if self.processed != None:
				self.write_daily_log(self.processed)
				self.upload_to_ftrack(self.processed)
				self.delete_log(self.log)
				self.close_this_app()
			else:
				self.delete_log(self.log)
				self.close_this_app()
		else:
			self.delete_log(self.log)
			self.close_this_app()

		#this doesnt seem to work, perhaps its because it cant delete its own log since it may be in use? run this through logger
		'''
		try:
			log_file = 'L:/HAL/Work_Tracker/dist/process_log.exe.log'
			statinfo = os.stat(log_file)
			if statinfo.st_size > 500000:
				self.delete_log(log_file)
		except:
			pass
		'''


	### runs ftrack uploader, creates daily log, deletes slog and closes itself. 

	##collects all the lines from the slog, and collapses by list of same keys
	def collapse_lines(self, _lst_of_dicts, _keys):
		_temp_lines = []

		for line in _lst_of_dicts:

			for index, tline in enumerate(_temp_lines):
				_keymatches = [key for key in _keys if line[key] == tline[key]]
				if len(_keys) == len(_keymatches):

					_temp_lines[index]['TaskDuration'] = tline['TaskDuration']+line['TaskDuration']

					break
			else:

				_temp_lines.append(line)

		return _temp_lines

	## contribute the unknown items to matching asset
	def contribute_unknown_task(self, _lst_of_dicts):
		
		_unknown_time = 0.0
		_total_useful_time = 1.0
		_useful_indexes = []


		return _lst_of_dicts

	## contribute the unknown items into all others
	def contribute_unknown(self, _lst_of_dicts):
		
		_unknown_time = 0.0
		_total_useful_time = 1.0
		_useful_indexes = []

		for index, line in enumerate(_lst_of_dicts):

			line['TaskDuration'] = 0.0

			if (line['Task'] == 'Unknown') and (line['Asset'] == 'Unknown'):
				#if its an unknown task, add dur to total, this total will be used for contribution
				_unknown_time += line['Duration']

			elif line['Task'] == 'Idle':
				#if its an idle task, dur = taskdur since values dont get contributed
				line['TaskDuration'] = line['Duration']
				pass

			else:
				#if its not idle add total useful time and append line index
				_total_useful_time += line['Duration']

				_useful_indexes.append(index)

		for i in _useful_indexes:

			dur = _lst_of_dicts[i]['Duration']
			#this converts the duration to task duration which has unknown time contribution to it.
			dur_factor = dur/_total_useful_time
			addition = dur_factor*_unknown_time

			_lst_of_dicts[i]['TaskDuration'] = dur + addition

		#if the only item in the list is the unknown task and asset, then taskduration = duration.
		if len(_useful_indexes) < 1:
			_lst_of_dicts[0]['TaskDuration'] = _lst_of_dicts[0]['Duration']

		return _lst_of_dicts

	## deletes a file with error handling
	def delete_log(self, _file):
		try:
			os.remove(_file)
			#pass
		except OSError:
			print( "..couldn't delete log! {0}".format(_file))
			pass
		return

	def upload_to_ftrack(self, _lst_of_dicts):
		os.environ['FTRACK_SERVER'] = 'https://domain.ftrackapp.com'
		os.environ['FTRACK_API_USER'] = os.environ.get("USERNAME")
		os.environ['FTRACK_API_KEY'] = '****************************************'
		session = ftrack_api.Session()

		items = _lst_of_dicts

		items = [item for item in items if item['Task'] and item['Asset'] not in ["Unknown","Idle"] ] 

		

		for item in items:
			if self.verbose:
				print (item)
			dur = item['TaskDuration']
			user = item['User']
			task = item['Task']
			asset = item['Asset']

			if dur > 20: #any recorded time less than this will not get updated to ftrack.
			

				available_tasks = session.query(
					'select id from Task '
					'where (parent.name is {1}) '
					'and (assignments any (resource.username = "{0}"))'
					.format(user, asset)
				)

				available_tasks = [str(i['name']) for i in available_tasks.all()] #['Precomp', 'Paintout', 'Retime', 'Compositing']

				if 'Compositing' in available_tasks:
					available_tasks.remove('Compositing')
					available_tasks.append('Compositing')

				if 'Matchmove' in available_tasks:
					available_tasks.remove('Matchmove')
					available_tasks.append('Matchmove')
					
				if task not in available_tasks:
					try:
						task = available_tasks[-1]
					except IndexError:
						pass

				current_task = session.query(
					'select id from Task '
					'where (parent.name is {1} and name is {2}) '
					'and (assignments any (resource.username = "{0}"))'
					.format(user, asset, task)
				)
				

				users = session.query( 'select username from User where username is {0}'.format(user) )
				current_user_id =  users.first()['id']
				#print (dur, task, asset)
				try:
					for i in current_task.all():
						new_timelog = session.create('Timelog', {'duration':dur, 'user_id':current_user_id})
						i['timelogs'].append(new_timelog)
				except ftrack_api.exception.ServerError:
					pass
				session.commit()
				#print (item)


	def write_daily_log(self, _lst_of_dicts):

		_date = datetime.now()
		_date = _date.strftime('%Y_%m_%d')


		daily_log_name = '{0}_{1}.txt'.format(self.user, _date)
		daily_log_dir = "{0}{1}/".format("L:/HAL/LIVEAPPS/utils/WorkTracker/_collections/", self.user)

		if not os.path.exists(daily_log_dir):
			os.makedirs(daily_log_dir)

		daily_log_dir = daily_log_dir+daily_log_name

		daily_log = '\n'+'\n'.join([str(i) for i in _lst_of_dicts])+'\n'
		#########print (daily_log)

		with open(daily_log_dir, "a+") as file:
			file.write(daily_log)

		return

	def read_slog(self):
		try:
			with open(self.log, "r+") as file:
				lines = file.readlines()

			lines = [eval(line) for line in lines if line != "\n"]
			
			lines = self.collapse_lines(lines, ['Task', 'Asset'])

			for i, line in enumerate(lines):
				lines[i]['Duration'] = lines[i]['TaskDuration']
				lines[i]['TaskDuration'] = 0.0
			

			lines = self.contribute_unknown(lines)
			return lines
		except IOError:
			#print ("closing....")
			self.close_this_app()

			return None

	def close_this_app(self):
		wmi = win32com.client.GetObject('winmgmts:')
		process_logs = [int(p.Properties_('ProcessId')) for p in wmi.InstancesOf('win32_process') if 'process_log.exe' in p.Name.lower()]
		for p in process_logs:

			PROCESS_TERMINATE = 1
			handle = win32api.OpenProcess(PROCESS_TERMINATE, False, p)
			win32api.TerminateProcess(handle, -1)
			win32api.CloseHandle(handle)
		return

process_log()