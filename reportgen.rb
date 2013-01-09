# reportgen.rb
# Ben Garvey
# bengarvey@garvey.com
# @bengarvey
# 01/09/2013
# Generates a list of projects to update

require_relative 'mspmg.rb'

m = MachineGun.new
m.dir = 'C:\Users\bengarvey\Dropbox\projects\msp_machinegun\projects'
#m.dir = 'U:\PROJ2000'
m.writeReport
  