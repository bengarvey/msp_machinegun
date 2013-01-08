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
 

 

 
 l = Listgen.new
 l.dir = 'C:\Users\bengarvey\Dropbox\projects\msproject\projects'
 #report = l.getCriticalTasks
 l.fire
 
 #File.open("report.txt", 'w') do |file|
#	file.puts(report)
 #end
 
 


