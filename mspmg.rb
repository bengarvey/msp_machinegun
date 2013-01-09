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
  
  def writeList
  
	list = ""
	files = getList	
	files.each do |f|
		list += /.*\/(.*\.mpp)/.match(f)[1] + "\n"
	end
	
	puts "Here's the list:  #{list}"
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
			puts "Setting #{tid} to #{actions[key][tid]} in #{key} (should be #{rem})"
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
  
  # Loops through projects and allows you to input new duractions for doable tasks
  def fire
   
	files = getList()  
	app = WIN32OLE.new("MSProject.Application")
	app.Visible = false	
	tid = ""
	donedone = false
	
	files.each do |f|	 
		
		app.FileOpen(f)	
		
		 # Use this to suppress "Are you sure?" messages		
		app.DisplayAlerts = false		
		project = app.ActiveProject	 
		tasks = project.Tasks
		done 		= false
		
		while !done && !donedone
		
			# Get project name and number
			puts "#{tasks[1].Name}:  Type task id/n/q/help"
			tid = gets.chomp
					
			# Checking input
			case tid 
				when 'q' then puts "All done!"			
					donedone = true
				when 'n', '', ' ' then puts "Next project..."
					#puts "Save changes before moving to next project? y/q"
						#response = gets.chomp
						#case response
							#when 'y' 
								puts "Advancing project..."
								
								d = Date.today
								puts d.strftime("%D") + " 5:00 PM"
								app.UpdateProject(true, d.strftime("%D") + " 5:00 PM", 2)
								
								puts "Saving..."
								app.FileSave
								puts "File saved"
								
								done = true
							#when 'n' 
							#	puts "Save skipped"
							#	done = true
							when 'q' 
								puts "Quitting"
								done = true
								donedone = true
						#end

				when 'help'
					puts "Type the task of a task to change it's remaining duration.\nType n to move onto the next project.\nType q to quit the process."					
				else 
					if tid.to_i > tasks.count
						puts "Couldn't find that task."
					else				
						puts "Current duration is #{tasks[tid].RemainingDuration.to_i / 480}, enter new duration: "					
						newduration = gets.chomp
							if newduration.to_i >= 0
								puts "Changing duration from #{tasks[tid].RemainingDuration.to_i / 480} to #{newduration.to_i}"
								tasks[tid].ActualDuration = tasks[tid].ActualDuration + 480
								tasks[tid].RemainingDuration = newduration.to_i * 480
								#puts "Check:  New rem duration is #{tasks[tid].RemainingDuration.to_i / 480}"
								#puts "Check:  New act duration is #{tasks[tid].ActualDuration.to_i / 480}"
							else
								puts "New values can't be less than zero"
							end
					end
			end
			
		end
		
=begin
		tasks.each do |t|		
			# Is this task doable and critical?
			if doable(t)	
				puts "#{f}\t#{t.Id}\t#{t.Name}\t#{t.RemainingDuration}\n"			
				report += "#{f}\t#{t.Id}\t#{t.Name}\t#{t.RemainingDuration}\n"
			end
		 end
=end		 
				 
		 app.FileClose
	end
	
	app.Quit 	
	
  end
  
  # Accepts a task and returns whether we can do it now or not
  def doable(task)
  		result = false
		
		if task.PercentComplete < 100 && task.FinishSlack == 0 && task.ResourceNames != ""
			if (Date.today + 1) >= Date.parse(task.EarlyStart.to_s[0,10])
				result = true
			end
		end
		
		return result
  end
    
 end
 

 

 
 #m = MachineGun.new
 #m.dir = 'C:\Users\bengarvey\Dropbox\projects\msp_machinegun\projects'
 #m.dir = 'U:\PROJ2000'
 
 #m.writeReport
 #m._fire
 
=begin
 File.open('list.txt').each_line { |s|
	puts s
 }
=end
 
=begin
 File.open('haggis.txt').each_line{ |s|
  puts s
}
=end
 


