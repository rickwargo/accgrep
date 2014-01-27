#!/usr/bin/ruby
require 'win32ole'
require File.dirname(__FILE__) + '/access_object'
require File.dirname(__FILE__) + '/option_parser'

require 'access_object'
exit if defined?(Ocra)

class Access 
	@@access = nil
	
	def initialize
		startup
	end
	
	def application
		@@access
	end
	
	def startup
		@@access = @@access.nil? ? WIN32OLE.new("Access.Application") : @@access
		@@access.UserControl = false
		@@access.Visible = false
		@@access.AutomationSecurity = 3 # msoAutomationSecurityForceDisable
		@@access.Echo false 
	end
	
	def finish
		@@access.Quit
		sleep 0.25
		@@access = nil
	end
	
	def recycle
		finish
		startup
	end
end

def grep1(access, file)
  matches = 0

  if Options.recurse && FileTest.directory?(file)
    grep(access, Dir["#{file.gsub(/\\/, "/")}/*.acc*"])
  end
		
  if FileTest.file?(file) && !FileTest.directory?(file) && file.match(Options.include) && (Options.exclude.empty? || !file.match(Options.exclude))
	access.OpenCurrentDatabase((file =~ /[\\\/]/ ? '' : Dir.getwd + '/') + file)
	$stderr.puts "Problem opening #{file}!" if access.nil?
	access.DoCmd.SetWarnings false	# does not seem to work
	access.DoCmd.Echo false					# does not seem to work
	access.Echo false								# does not seem to work

	begin
		needs_save = 0
		Options.search.each do |object_type|
			modifier = Options.controls.empty? ? Options.properties.empty? ? "" : " properties" : " controls"
			$stderr.puts "Searching #{object_type.sub(/#{modifier.empty? ? "" : "s$"}/, "")}#{modifier} in #{file} for \"#{Options.expression.join("\"|\"")}\"" if Options.verbose
			access_obj = AccessObject.new(object_type, access)
			access_obj.proc = Options.procedure if object_type =~ /^proc/
			access_obj.each do |obj|
				Options.expression.each do |expr|
				opts = 0
				opts |= Regexp::IGNORECASE if Options.ignore_case
				opts |= Regexp::EXTENDED if Options.extended
				opts |= Regexp::MULTILINE if Options.multi_line
				re = Regexp.new(expr, opts)
					if re.match(obj) || Options.invert_match
						if Options.files_with_matches 
							puts file
							return
						elsif Options.delete_matching_line
							$stderr.puts "Deleting line matching #{re} in #{file}" if Options.verbose
							access_obj.delete_current_line
							needs_save = needs_save + 1
						elsif !Options.replace.empty?
							t = obj.sub(re, Options.replace)
							if t != obj
								$stderr.puts "Replacing #{re} in #{file} with \"#{Options.replace}\"" if Options.verbose
								access_obj.replace(t)
								needs_save = needs_save + 1
							end
						else
							printf("%s%s%s\n",
								Options.recurse || ARGV.length > 1 ? "[#{file}] " : "",
								Options.line_numbers ? "#{access_obj.where}: " : "",
								obj)
                
							matches += 1
							return if Options.max_count.to_i > 0 && matches >= Options.max_count.to_i
						end
					end
				end
			end
		end if !access.nil?
	ensure
		access.Save if !access.nil? && needs_save > 0
		access.CloseCurrentDatabase if !access.nil?
		access = nil
	end
	if Options.files_without_matches && matches == 0
		puts file
	end
  end  
end

def grep(access, files)
  n = 0
  files.each do |file|
    grep1(access.application, file)
		access.recycle if Options.recycle_every > 0 && (n+=1) % Options.recycle_every == 0
  end
end

begin # if __FILE__ == $0
  Options = OptionParser.parse(ARGV)
  
	access = Access.new
	begin
		grep(access, ARGV)
	ensure
		access.finish
	end
end
