# mspmg.rb
# Ben Garvey
# bengarvey@garvey.com
# @bengarvey
# 01/02/2013
# Looks through a directory of MS Project files and generates a lit of thoe projects

require 'rubygems'
require 'win32ole'
require 'date'

class MachineGun
  
  attr_accessor :dir, :files
  
  def initialize
	@files = Array.new
  end
  
 def fire 
  	
	# Initialize some variables
	tid 		= ""
	donedone 	= false  
	actions 	= Hash.new
	key 		= ""
  
	# Get our and create actions
	File.open('list.txt').each_line { |s|
		if /.*\.mpp/.match(s) 
			key = /.*\.mpp/.match(s)[0]
			actions[key] = Hash.new
		elsif /(\d*),(\d*)/.match(s)			
			tid = /(\d*),(\d*)/.match(s)[1]
			rem = /(\d*),(\d*)/.match(s)[2]
			actions[key][tid] = rem
			puts "Setting #{tid} to #{actions[key][tid]} in #{key}"
		end
	} 
	
	puts
	
	# Open MS Project
	app = WIN32OLE.new("MSProject.Application")
	
	# Run in the background
	app.Visible = false	
	
	# Loop through hash and open each file
	actions.each_key do |k|
		
		# Open MS Project file
		app.FileOpen("#{@dir}\\#{k}")	
		
		# Use this to suppress "Are you sure?" messages		
		app.DisplayAlerts = false
		
		# Initialize references
		project = app.ActiveProject	 
		tasks 	= project.Tasks
	
		puts "Opening #{k}"
		puts "Total tasks: #{tasks.count}"
		
		# Loop through each action
		actions[k].each_key do |t|
			# Make sure this is a valid task
			if t.to_i < tasks.count && t.to_i > 0
				puts "\tSetting #{t} to #{actions[k][t]} for #{k}"
				puts "\tIncreasing #{t}'s actual duration to #{ (tasks[t].ActualDuration + 480) / 480} for #{k}"
				tasks[t].ActualDuration 	= tasks[t].ActualDuration + 480
				tasks[t].RemainingDuration 	= actions[k][t].to_i * 480
			else
				puts "Couldn't find task #{t} in #{k}. Skipping it!"
			end
		end
		
		# Schedule new work after today
		puts "Advancing project..."		
		d = Date.today
		dstr = d.strftime("%D") + " 5:00 PM"
		puts "Scheduling tasks to start after #{dstr}"
		app.UpdateProject(true, dstr, 2)
		
		# Save file and close
		puts "Saving..."
		app.FileSave
		
		puts "File saved\n\n"		
		app.FileClose
		
	end
	
	puts "All files updated";
		
	app.Quit 	
	
  end
  
  def printList
	Dir.entries(@dir).each do |p|
		if /.\.mpp/.match(p) && !/master\.mpp/.match(p)
			puts p
			@files.push("#{@dir}/#{p}")
		end
	end
  end
  
  def getList
  
 	Dir.entries(@dir).each do |p|
		if /.\.mpp/.match(p) && !/master\.mpp/.match(p)
			@files.push("#{@dir}/#{p}")
		end
	end
	
	return @files
	
  end
  
  def writeReport
  
	list = ""
	files = getList	
	files.each do |f|
		list += /.*\/(.*\.mpp)/.match(f)[1] + "\n"
	end
	
	puts "Here's the list:  \n#{list}"
	file = File.open('list.txt', 'w') { |file| file.write(list) }
	
  end
  
  def getCriticalTasks
  
	report = ""  
	files = getList()  
	app = WIN32OLE.new("MSProject.Application")
	app.Visible = false	
  
	files.each do |f|	 
	
		app.FileOpen(f)	 
		project = app.ActiveProject	 
		tasks = project.Tasks
			 
		# Loop through all tasks
		tasks.each do |t|		
			# Is this task doable and critical?
			if doable(t)	
				puts "#{f}\t#{t.Id}\t#{t.Name}\t#{t.RemainingDuration}\n"			
				report += "#{f}\t#{t.Id}\t#{t.Name}\t#{t.RemainingDuration}\n"
			end
		 end
		 
		 app.FileClose
	end
	
	app.Quit 
	
	return report
	
  end
  
   class String
    def is_i?
       !!(self =~ /^[-+]?[0-9]+$/)
    end
   end
    
 end
 


