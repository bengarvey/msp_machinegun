msp_machinegun
==============

Quickly update and schedule new work on a directory of MS Project files

We were tediously going through and edited MS Project Files, so this script opens up all the files in a directory and asks for updates for each one.  It then reschedules all new work to be started the next day.

If anything, it's a good reference for how to manipulate MS Project files in Ruby.

To use it, set the paths in reportgen.rb and update.rb to where your project files are.  A file called list.txt is created.

There, you can make simple edits to list.txt that will later make changes in the MS Project file.  For example, if your generated output is

+sch1.mpp
+sch2.mpp
+sch3.mpp
+sch4.mpp

You can edit it so it looks like this

+sch1.mpp
+20,0
+21,1
+22,0
+ch2.mpp
+15,1
+sch3.mpp
+sch4.mpp
+90,0
+75,5
+30,1

Save list.txt and then run update.rb.  The above will make the following changes.
+sch1.mpp
+20,0 # 20th task will have 0 remaining days and its actual duration will go up by 1
+21, 1 # 21st task will have 1 remaining day and it's actual duration will go up by 1
+sch2.mpp
+15, 1 # sch2.mpp's 15th task will have 1 remaining day and it's actual duraction will be +1
+sch3.mpp # No change
+ch4.mpp
+90,0 # Same as the others above.  You get the idea
+5,5,
+30,1

