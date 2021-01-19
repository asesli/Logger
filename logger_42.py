import collections
from datetime import datetime, timedelta
import getpass
import itertools
import os
import sys
sys.coinit_flags = 0 
##PSUTIL ERRORS OUT ON WIN10

import psutil
import win32com.client

import pythoncom
import re
from shutil import copyfile
from socket import gethostname
import threading
#import thread

#import win32api
#import win32con
import win32gui
import win32process
#import win32security

#import anydbm
#import dbhash

import gc


class logger():
	def __init__(self):

		#Kill current proc if logger.exe is already running.
		##PSUTIL ERRORS OUT ON WIN10

		
		if len([p.name() for p in psutil.process_iter() if p.name() == 'logger.exe']) > 1:
			sys.exit()
		
		'''
		wmi = win32com.client.GetObject('winmgmts:')
		if len([p for p in wmi.InstancesOf('win32_process') if 'logger' in p.Name.lower()]) > 1:
			sys.exit()
		'''


		self.collected_apps = list() 
		self.user = getpass.getuser()
		self.user_dir = os.environ['USERPROFILE']+'\\'

		self.log = "{0}_log_{1}.txt".format(self.user_dir, self.user)
		self.server_log = "{0}log_{1}_{2}.txt".format("L:/HAL/LIVEAPPS/utils/WorkTracker/_slog/", self.user, gethostname())

		self.delete_log(self.log)
		self.first_run = True
		self.last_line={}

		#self.process_log_file = '"L:\\HAL\\LIVEAPPS\\utils\\WorkTracker\\bin\\dist\\process_log.exe"'
		self.process_log_file = '"L:\\HAL\\LIVEAPPS\\utils\\WorkTracker\\bin\\process_log\\process_log.exe"'
		#self.process_log_file = '"C:\\Users\\luxali\\output\\process_log_20\\process_log_20.exe"'

		#verbose for memory_log, raw_log, server_log
		self.verbose = True, True, True
		#master verbose switch
		self.printout = False

		if not self.printout: 
			self.verbose = False, False, False

		self.gc = False

		self.settings = 0

		if self.settings == 0:
			#production settings
			self.memory_i = 5
			self.rlog_i = 30
			self.idle_i = 600
			self.slog_i = 602
			self.flog_i = 1810
			
		if self.settings == 1:
			#debug settings
			self.memory_i = 1
			self.rlog_i = 5
			self.idle_i = 600
			self.slog_i = 31
			self.flog_i = 66
			
		if self.settings == 2:
			#aggressive testing for memory leaks
			self.memory_i = 0.2
			self.rlog_i = 1
			self.idle_i = 600
			self.slog_i = 5
			self.flog_i = 15
		
		
		self.update_ftrack_timelog()
		self.collect_apps_wins()
		self.save_to_raw_log()
		self.save_to_server_log()


	
	## get_active_windows returns a list of window names connected to the active window.
	def get_active_windows(self):

		def _get_associated_windows(_tid, _pid):
			"""return all windows associated to pid and tid. output is converted to string separated by spaces."""

			def _enum_all_windows():
				def callback(handle, data):
					container_list = []
					container_list.append(win32gui.GetWindowText(handle))
					_tid, _pid = win32process.GetWindowThreadProcessId(handle)
					container_list.append(_tid)
					container_list.append(_pid)
					titles.append(container_list)
				titles = []
				win32gui.EnumWindows(callback, None)
				if self.gc:
					gc.collect()
				return titles

			all_windows = _enum_all_windows()

			_associated_proc = set()
			for i in all_windows:
				if (i[2] == _pid) and (i[1] == _tid) and (i[0]!='') and (i[0]!=' ') and (i[0]!='Default IME') and (i[0]!='MSCTFIME UI') :

					_associated_proc.add(i[0])

			_associated_proc = list(_associated_proc)
			if self.gc:
				gc.collect()
			return _associated_proc

		x = win32gui
		x = x.GetForegroundWindow()
		tid,pid = win32process.GetWindowThreadProcessId(x)
		associated_proc = _get_associated_windows(tid,pid)
		
		associated_proc = " ".join(associated_proc)

		#mayas produces invisible windows based on operation. it is usually separated by '   ---   '
		associated_proc = associated_proc.split('   ---   ')[0]
		if self.verbose[0]:
			print( 'Memory Log Line: ',associated_proc)######################
		if self.gc:
			gc.collect()
		return associated_proc
		

	## returns timestamp (str)
	def get_date(self):
		"""returns the current timestamp"""
		_date = datetime.now()
		_date = _date.strftime('%Y/%m/%d %H:%M:%S')
		return str(_date)

	## deletes a file with error handling
	def delete_log(self, _file):
		try:
			os.remove(_file)
		except OSError:
			print( "..couldn't delete log! {0}".format(_file))
			pass
		return

	## collapses same WInfo and Apps, and add the durations
	def collapse_lines(self, _lst_of_dicts, _keys):
		_temp_lines = []

		for line in _lst_of_dicts:

			for index, tline in enumerate(_temp_lines):
				_keymatches = [key for key in _keys if line[key] == tline[key]]
				if len(_keys) == len(_keymatches):

					_temp_lines[index]['Duration'] = tline['Duration']+line['Duration']
					if 'TaskDuration' in _temp_lines[index]:
						_temp_lines[index]['TaskDuration'] = tline['TaskDuration']+line['TaskDuration']
					break
			else:

				_temp_lines.append(line)

		return _temp_lines
	
	## contribute the unknown items into dicts
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
		try:
			if len(_useful_indexes) < 1:
				_lst_of_dicts[0]['TaskDuration'] = _lst_of_dicts[0]['Duration']
		except IndexError:
			pass

		return _lst_of_dicts

	## 
	## main functions

	## writes memory log as raw log
	def write_rlog_line(self):

		## sorts a list of strings by occurance (Most to Least) without duplicates (list) made up of (str)s
		def _sort_by_most_occured(_lst_of_strs):
			'''sorts a list of strings based on occurance and removes duplicates; returns list of strings'''
			counts = collections.Counter(_lst_of_strs)
			_lst_of_strs = sorted(_lst_of_strs, key=counts.get, reverse=True)

			seen = set()
			seen_add = seen.add
			_lst_of_strs = [x for x in _lst_of_strs if not (x in seen or seen_add(x))]
			_lst_of_strs = list(_lst_of_strs)
			return _lst_of_strs

		## returns timestampA-timestampB in seconds (float)
		def _get_duration(str_timestamp):
			"""Get duration in seconds by the string param. param requires strict naming rules"""
			last_stamp = str_timestamp
			last_stamp = re.findall("[-+]?\d+[\.]?\d*[eE]?[-+]?\d*", last_stamp)

			for  i, val in enumerate(last_stamp):
				last_stamp[i] = int(val)
			l = last_stamp

			last_stamp = datetime(l[0], l[1], l[2], l[3], l[4], l[5])
			_duration =  datetime.now()-last_stamp
			_duration = _duration.total_seconds()
			if self.gc:
				gc.collect()
			return float(_duration)

		## returns mouse pos x,y (tuple)
		def _get_mouse_pos():
			"""return mouse position x,y as tuple"""
			_mousePosition = (0,0)
			try:
				_mousePosition = win32api.GetCursorPos()

			except:
				_mousePosition = (0,0)
			if self.gc:
				gc.collect()
			return _mousePosition

		hwnd = win32gui
		hwnd = hwnd.GetForegroundWindow()

		#Current

		date = self.get_date()
		user = self.user
		mousePosition = _get_mouse_pos()
		tid,pid = win32process.GetWindowThreadProcessId(hwnd)
		total_collected_apps = _sort_by_most_occured(self.collected_apps)
		duration = 0.0
		last_rlog_line = {}
		log = self.log

		new_rlog_line = {
		'Date': date,
		'User': user,
		'WInfo': {'CursorXY': mousePosition,'Tid': tid,'Pid': pid},
		'Apps': total_collected_apps,
		'Duration': duration
		}
		#print( str(new_rlog_line))
		
		w_err = False
		try:
			with open(log, "r+") as file:
				try:
					lines = file.read().splitlines()
					baked_lines = lines[:-1]
					last_line = lines[-1]
					last_line = eval(last_line)
					last_duration = _get_duration(last_line['Date'])
					last_line['Duration'] = last_duration
					lines = baked_lines
					lines.append(last_line)
					lines.append(new_rlog_line)
				except IndexError:
					w_err = True
		except IOError:
			w_err = True

		with open(log, "w+") as file:
			if w_err:
				lines = [new_rlog_line]

			lines = '\n'.join([str(i) for i in lines])
			file.write(lines)
			if self.verbose[1]:
				print  ('\nRaw Log Line: ',lines,'\n')########################
		del self.collected_apps[:]

	## writes raw log to server log
	def write_slog_line(self):

		file = self.log

		with open(file, "r+") as _logfile:
			lines = _logfile.readlines()
		lines2 = []
		for i in lines:
			try:
				if i != '\n':
					i = eval(i)
					if i['Duration'] == 0.0:
						i['Duration'] = float(self.rlog_i)#force set duration of the last item. 
					lines2.append(i)

			except SyntaxError:
				print ('problem while evaluating line from self.log')
				pass
		lines = lines2

		## creates a new list of dicts w/ date,user,dur,task,asset
		def _convert_rlog_to_slog(_lst_of_dicts):
			_temp_lines = []
			#full_pattern:  returns full retrieveble info from the windows. from here you can narrow down shot code and task name.. BEY209_029_040_Cleanup_V036.nk - , BEY209_029_040_V036.nk - 
			full_pattern = re.compile(r"([a-z]{2,3}(\d{3})*_\d{3}_\d{3}.*\\(anim|light|lighting|fx|trac|tracking|layout|setup)\\)|(\\05_comp\\.*[a-z]{2,3}(\d{3})*_\d{3}_\d{3}).*|([a-z]{2,3}(\d{3})*_\d{3}_\d{3}[^ ]*((\.nk)|(\.autosave))\W* )|(T:\\01_Assets\\.*\\[a-z]{2,3}_\w*\\)|(PFTrack.*[a-z]{2,3}(\d{3})*_\d{3}_\d{3}([^\_\W])*)", re.I)#|([a-z]{3}(\d{3})*_\d{3}_\d{3})", re.I)
			#shotcode(asset) pattern: return a shot code.. BEY201_010_025, WIL_010_40, etc
			shotcode_pattern = re.compile(r"([a-z]{2,3}(\d{3})*_\d{3}_\d{3}([^\_\W])*)", re.I)
			#print ('#############')

			for line in _lst_of_dicts:
				apps = line['Apps']
				## find asset and task, if it cannot find both, search next app from app list. 
				for app in apps:
					app = app.replace('/', '\\')
					#print( app)
					
					task = 'Unknown'
					asset = 'Unknown'

					#find full pattern
					f_match = full_pattern.search(str(app))
					#print (f_match)
					try:
						match = f_match.group()

						if match:
							#print (match)

							a_match = shotcode_pattern.search(str(app))
							a_match = a_match.group()
							asset = a_match

							if '\\anim\\' in match: task = 'Animation'
							elif '\\trac\\' in match: task = 'Tracking'
							elif '\\tracking\\' in match: task = 'Tracking'
							elif '\\light\\' in match: task = 'Lighting'
							elif '\\lighting\\' in match: task = 'Lighting'
							elif '\\fx\\' in match: task = 'FX'
							elif '\\layout\\' in match: task = 'Layout'
							elif '\\setup\\' in match: task = 'Tracking'
							#elif '\\05_COMP\\' in match: task = 'Compositing'#what about when opening a precomp sequence in dj_view? currently it will guess a comp task... THATS FINE. precomps are recorded only through nuke.
							elif bool(re.search(r'(([a-zA-Z]{2,3}(\d{3})*_\d{3}_\d{3})[^ ]*(\.nk))', match)): 

								task = 'Compositing'

								_r0 = re.compile(r"([a-z]{2,3}(\d{3})*_\d{3}_\d{3}([^\_\W])*)_", re.I)
								_r1 = re.compile(r"_?v(\d{3}).*", re.I)

								_r0_match = _r0.search(str(match))
								_r0_match = _r0_match.group()

								_r1_match = _r1.search(str(match.replace(_r0_match, '')))
								_r1_match = _r1_match.group()

								task = match.replace(_r0_match, '').replace(_r1_match, '')

								if task.lower() in ['comp', '']:

									task = 'Compositing'

								#if '01_precomp' in task.lower():

								'''
								_rt0 = re.compile(r"\W", re.I)
								_rt0_match = _rt0.search(str(task.lower()))
								_rt0_match = _rt0_match.group()

								print (_rt0_match)
								'''


							elif bool(re.search(r'(PFTrack.*[a-zA-Z]{2,3}(\d{3})*_\d{3}_\d{3})', match)):
								task = 'Tracking'

							else:

								if asset != match:
									task = match


								_rt0 = re.compile(r"\W", re.I)
								_rt0_match = _rt0.search(str(task))
								_rt0_match = _rt0_match.group()

								if _rt0_match != None:
									if '\\05_COMP\\' in task:
										if self.printout:
											print ("could not guess")
										###this is great and all.... but waht if im looking at a random precomp in djview? PrecompABC may not always be an assigned task. maybe create a new dept?
										'''
										if bool(re.search(r".*\\01_precomp\\", match)):
											_rt1 = re.compile(r".*\\01_precomp\\", re.I)
											_rt1_match = _rt1.search(str(task))
											_rt1_match = _rt1_match.group()
											#print _rt1_match
											task = task.split(_rt1_match)[-1]
											task = task.split('\\')[0]
										else:
											task = 'Compositing'
										'''
										task = 'Compositing'


					except AttributeError:
						pass
					if self.printout:
						print (task, asset, '\n')
					line['Task'] = task
					line['Asset'] = asset
					
					if (task != 'Unknown') and (asset != 'Unknown'):
						break

				if line['Duration'] > self.idle_i:
					line['Task'] = 'Idle'
					line['Asset'] = 'Idle'


				del line['WInfo'], line['Apps']
				#print (line)
				#print ('')
				_temp_lines.append(line)
				#print (info_str)

			return _temp_lines

		lines = self.collapse_lines(lines, ['WInfo','Apps'])
		lines = _convert_rlog_to_slog(lines)
		lines = self.collapse_lines(lines, ['Task','Asset'])
		lines = self.contribute_unknown(lines)

		#save lines into a server log. 
		with open(self.server_log, "a+") as file:
			lines = '\n'.join([str(i) for i in lines])+'\n'
			
			file.write(lines)

		if self.verbose[2]:
			print ('Server Log Line: ',lines)#########################
		#print lines
		#self.delete_log(self.log)

	## launches the process log script to upload to ftrack
	def launch_process_log(self):

		'''
		try:
			log_file = 'L:/HAL/LIVEAPPS/utils/WorkTracker/bin/dist/process_log.exe.log'
			statinfo = os.stat(log_file)
			if statinfo.st_size > 500000:
				self.delete_log(log_file)
		except:
			pass
		'''
		##PSUTIL ERRORS OUT ON WIN10
		process_log_running = ["process_log.exe" in (p.name() for p in psutil.process_iter())][0]
		'''
		pythoncom.CoInitialize()
		wmi = win32com.client.GetObject('winmgmts:')
		process_log_running = [p for p in wmi.InstancesOf('win32_process') if 'process_log.exe' in p.Name.lower()]
		'''

		if not process_log_running :
			if self.printout:
				print ('processing log..')
			try:
				#os.startfile('"L:\\HAL\\LIVEAPPS\\utils\\WorkTracker\\bin\\dist\\process_log.exe"')
				#os.startfile('"C:\\Users\\luxali\\output\\process_log_20\\process_log_20.exe"')
				os.startfile(self.process_log_file)
			except:
				print ('couldnt launch process_log.exe')
				pass
		pythoncom.CoUninitialize()
	

	## 
	## interval starter functions

	## every 1 second, append active windows and exe into a list in memory
	def collect_apps_wins(self):
		threading.Timer(self.memory_i, self.collect_apps_wins).start()
		active_apps = self.get_active_windows()
		self.collected_apps.append(active_apps)

	## every 10 seconds sort out the list in memory, gather other info, generate duration and append as a new line in a raw log file. 
	def save_to_raw_log(self):
		threading.Timer(self.rlog_i, self.save_to_raw_log).start()
		if not self.first_run:
			self.write_rlog_line()

	## every 60 seconds, look through the raw log, convert the log append the reuslt as a new item in a server log file.
	def save_to_server_log(self):
		threading.Timer(self.slog_i, self.save_to_server_log).start()
		if not self.first_run:
			self.write_slog_line()
			try:
				with open(self.log, "w+") as file:
					file.write('')
			except IOError:
				print( 'could not clear log.. {0}'.format(self.log))
				pass
		self.first_run = False

	## every hour, collapse the log, and upload it to ftrack through separate executable. 
	def update_ftrack_timelog(self):
		threading.Timer(self.flog_i, self.update_ftrack_timelog).start()
		#upload to ftrack . exe will look at this txt to upload parameters
		self.launch_process_log()

logger()