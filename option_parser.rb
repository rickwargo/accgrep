require 'optparse'
require 'ostruct'

class OptionParser
  def self.parse(args)
    @search = %w[all macros procedures references forms reports tables queries data]
    
    options = OpenStruct.new
    options.files_with_matches = false
    options.files_without_matches = false
    options.no_messages = false
    options.ignore_case = false
    options.extended = false
    options.multi_line = false
    options.invert_match = false
    options.line_numbers = false
    options.verbose = false
    options.recursive = false
    options.delete_matching_line = false
    options.max_count = -1
    options.expression = []
    options.replace = ""
    options.search = []
    options.procedures = []
    options.procedure = ""
    options.include = ""
    options.exclude = ""
    options.forms_matching = ""
    options.queries_matching = ""
    options.reports_matching = ""
    options.tables_matching = ""
		options.linked_tables = false
    options.controls = []
    options.properties = []
		options.fields = []
		options.where_clause = ""
		options.recycle_every = 0
    
    opts = OptionParser.new do |opt|
      opt.banner = "Usage: #{$0.sub(/.*[\\\/]/, '').sub(/\.[^.]*/, '')} [options] [expression] file ..."
      
      opt.separator ""
      opt.separator "Options:"
      
      opt.on("-l", "--files-with-matches", "only print file names containing matches") do
        options.files_with_matches = true
      end
      
      opt.on("-L", "--files-without-matches", "only print file names containing no match") do
        options.files_without_matches = true
      end
      
#     opt.on("-q", "--no-messages", "suppress error messages") do
#       options.files_with_matches = true
#     end
      
      opt.on("-i", "--ignore-case", "ignore case distinctions") do
        options.ignore_case = true
      end
      
      opt.on("-X", "--extended", "use extended regular expressions") do
        options.extended = true
      end
      
      opt.on("-M", "--multi-line", "search across lines") do
        options.multi_line = true
      end
      
      opt.on("-v", "--invert-match", "select non-matching lines") do
        options.invert_match = true
      end
      
      opt.on("-n", "--line-numbers", "print line number with output lines") do
        options.line_numbers = true
      end
      
      opt.on("--recurse", "recurse into directories") do
        options.recurse = true
      end
      
      opt.on("-e", "--regexp PATTERN", "use PATTERN as a regular expression (may have multiples)") do |pattern|
        options.expression << pattern
      end
  
      opt.on("-r", "--replace STRING", "use STRING as a replacement to the regular expression") do |str|
        str.gsub!(/&dq;/, '""')
        options.replace << str
      end
  
      opt.on("-D", "--delete-matching-line", "delete lines matching the regular expression") do
        options.delete_matching_line = true
      end
  
      code_list = @search.join(', ')
      opt.on("-s", "--search WHAT", @search, "database objects to search", "  (#{code_list})") do |what|
        if what == "all"
          @search.each do |w|
            options.search |= [w] if w !~ /^all|^data/
          end
        else
          options.search |= [what]
        end
      end
      
      opt.on("-c", "--controls NAME", "search only controls matching NAME (should include property)") do |path|
        options.controls |= [path]
      end

      opt.on("-p", "--properties NAME", "search only properties matching NAME (should include form or report name)") do |path|
        options.properties |= [path]
      end

      opt.on("-f", "--field NAME", "search only field named NAME") do |fieldname|
        options.fields |= ['['+fieldname+']']
      end

      opt.on("-P", "--procedure NAME", "search only the NAMEd procedure (may have multiples)") do |name|
        options.search |= ['procedures']
        options.procedures << name
        options.procedure = name
      end
  
      opt.on("-m", "--max-count NUM", "stop after NUM matches") do |num|
        options.max_count = num
      end
      
      opt.on("--include PATTERN", "files that match PATTERN will be examined") do |pattern|
        options.include = pattern
      end
    
      opt.on("--exclude PATTERN", "files that match PATTERN will be skipped") do |pattern|
        options.exclude = pattern
      end
    
      opt.on("-F", "--forms-matching PATTERN", "forms that match PATTERN will be examined") do |pattern|
        options.forms_matching = pattern
				options.search |= ['forms']
			end
    
      opt.on("-Q", "--queries-matching PATTERN", "queries that match PATTERN will be examined") do |pattern|
        options.queries_matching = pattern
				options.search |= ['queries']
				end

				opt.on("-R", "--reports-matching PATTERN", "reports that match PATTERN will be examined") do |pattern|
        options.reports_matching = pattern
				options.search |= ['reports']
      end

      opt.on("-T", "--tables-matching PATTERN", "tables that match PATTERN will be examined") do |pattern|
        options.tables_matching = pattern
				options.search |= ['tables'] if !options.search.include?('data')
			end

      opt.on("--linked-tables", "only search linked tables (Connect string is not empty)") do
        options.linked_tables = true
				options.search |= ['tables'] if !options.search.include?('data')
      end
      
      opt.on("-w", "--where CLAUSE", "use clause to limit rows in searching table data") do |clause|
        options.where_clause << " (#{clause})"
      end
			
			opt.on("-a", "--and", "specify AND between the where clauses") do
				options.where_clause << " and"
			end

			opt.on("-o", "--or", "specify OR between the where clauses") do
				options.where_clause << " or"
			end

      opt.separator ""
      opt.separator "Options that shouldn't be options:"
      
      opt.on("-C", "--recycle-every NUM", "recycle access application every NUM times") do |n|
        options.recycle_every = n.to_i
      end
      
      opt.separator ""
      opt.separator "Common options:"
      
      opt.on_tail("-h", "--help", "show this message") do
        puts opts
        exit!
      end
      
      opt.on_tail("-V", "--verbose", "show messages indicating progress") do
        options.verbose = true
      end
      
      opt.on_tail("--version", "show version") do
        puts OptionParser::Version.join('.')
        exit!
      end
    end
    
    begin
      opts.parse!(args)
      if options.expression.length == 0 && ARGV.length > 0 && !FileTest.file?(ARGV[0])
        options.expression = [ARGV[0]] 
        ARGV.shift
      end
      options.expression = [""] if options.expression.length == 0 # justs spits out all lines/cells
      options.search |= ["macros"] if options.search.length == 0
      bad_files = 0
      ARGV.each do |file|
        if !FileTest.readable?(file)
          $stderr.puts "#{file}: Cannot open #{file} for reading."
          bad_files += 1  
        end
      end
      raise StandardError.new, "Cannot open #{bad_files} file(s) for reading." if bad_files > 0
      raise StandardError.new, "Must specify at least one file." if ARGV.length == 0
    rescue StandardError => msg
      puts msg, ""
      puts opts
      exit!
    end
    options
  end
end
