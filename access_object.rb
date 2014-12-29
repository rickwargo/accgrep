class AccessObject
  class AbstractAccessObject
    def initialize(access)
      @access = access
      @form = nil
			@control = nil
			@property = ""
			@query = ""
			@report = ""
			@table = ""
			@field = ""
			@recno = 0
    end

    def each_form
			@access.CurrentProject.AllForms.each do |form|
				if form.Name.match(Options.forms_matching)
					$stderr.puts ">>> Searching form \"#{form.Name}\"" if Options.verbose
					@access.DoCmd.OpenForm form.Name, acViewDesign=1
					@form = @access.Screen.ActiveForm  # for Access
					yield @form 
					@access.DoCmd.Close acForm=2, form.Name, acSaveNo=2
				end
      end
    end
    
    def each_report
			@access.CurrentProject.AllReports.each do |report|
				if report.Name.match(Options.reports_matching)
					$stderr.puts ">>> Searching report \"#{report.Name}\"" if Options.verbose
					@access.DoCmd.OpenReport report.Name, acViewDesign=1
					@report = @access.Screen.ActiveReport  # for Access
					yield @report 
					@access.DoCmd.Close acReport=3, report.Name, acSaveNo=2
				end
			end
		end
		
    def each_query
			@access.CurrentDb.QueryDefs.each do |@query|
				if @query.Name.match(Options.queries_matching) && (@query.Connect.empty? || Options.linked_tables)
					$stderr.puts ">>> Searching query \"#{@query.Name}\"" if Options.verbose
					yield @query 
				end
      end
    end
    
    def each_table
			@access.CurrentDb.TableDefs.each do |@table|
				if @table.Name.match(Options.tables_matching) && (@table.Connect.empty? || Options.linked_tables)
					$stderr.puts ">>> Searching table \"#{@table.Name}\"" if Options.verbose
					yield @table 
				end
      end
    end
		
		def each_datum
			@access.CurrentDb.TableDefs.each do |@table|
				if @table.Name.match(Options.tables_matching) && (@table.Connect.empty? || Options.linked_tables)
					$stderr.puts ">>> Searching table data in \"#{@table.Name}\"" if Options.verbose
					yield @table 
				end
      end
		end
		
		def each_control
			raise NotImplementedError
		end
    
    def each
      raise NotImplementedError
    end
    
    def where_line
      "#{@line}"
    end
  end

  class ControlAccessObject < AbstractAccessObject
    def where
      "#{@name}"
		end
		
		def iterator(&block)
			each_control &block
		end
		
		def open
		end
		
		def close
		end
		
		def active_object
			@control
		end
    
		def each_control
			iterator do |object| # objects have already been filtered
				object.Controls.each do |control|
					@control = control
					yield control 
				end
			end
		end

		def each
			each_control do |control|
				control.Properties.each do |property|
					Options.controls.each do |control_property|
						name = "#{control.Name}.#{property.Name}" 
						if name =~ /#{control_property}/
							@name = name
							@control = control
							@property = property.Name
							val = nil
							begin
								val = @control[@property].to_s
							rescue
								val = nil
							end
              yield val unless val.nil?
						end
					end
				end
			end
		end

    def replace(new_value)
      @control[@property] = new_value
    end
  end
	
	class ReportControlAccessObject < ControlAccessObject
    def where
      "#{@report.Name}!#{@name}"
		end

		def active_object
			@access.Screen.ActiveReport
		end

		def iterator(&block)
			each_report &block
		end
	
		def open(report)
			@access.DoCmd.OpenReport report.Name, acViewDesign=1
			@report = active_object  # for Access
		end
		
		def close(report)
			@access.DoCmd.Close acReport=3, report.Name, acSaveNo=2
		end
	end
	
	class FormControlAccessObject < ControlAccessObject
    def where
      "#{@form.Name}!#{@name}"
		end
		
		def active_object
			@access.Screen.ActiveForm
		end
		
		def iterator(&block)
			each_form &block
		end
		
		def open(form)
			@access.DoCmd.OpenForm form.Name, acViewDesign=1
			@form = active_object  # for Access
		end
		
		def close(form)
			@access.DoCmd.Close acForm=2, form.Name, acSaveNo=2
		end	
  end

	class LineAccessObject < AbstractAccessObject
    def where
      where_line
    end

    def iterate(objects, &block)
      @line = 0
      objects.each do |obj|
        @line += 1
        begin
          val = obj.Value.to_s
        rescue # in case COM fails us
          val = "?"
        end
        yield obj.Name + ": " + val
      end
    end
  end

	class PropertyAccessObject < AbstractAccessObject
		def where
			"#{@object.Name}.#{@property}"
		end

		def each(&block)
			iterator do |object|
				object.Properties.each do |property|
					Options.properties.each do |object_property|
						name = "#{object.Name}.#{property.Name}" 
						if name =~ /#{object_property}/
							@name = name
							@object = object
							@property = property.Name
							val = nil
							begin
								val = object[@property].to_s
							rescue
								val = nil
							end
              yield val unless val.nil?
						end
					end
				end
			end
		end
	end
	
	class FormPropertyAccessObject < PropertyAccessObject
		def iterator(&block)
			each_form &block
		end
	end
	
	class ReportPropertyAccessObject < PropertyAccessObject
		def iterator(&block)
			each_report &block
		end
	end
	
	class TablePropertyAccessObject < PropertyAccessObject
		def iterator(&block)
			each_table &block
		end
	end
	
	class QueryPropertyAccessObject < PropertyAccessObject
		def iterator(&block)
			each_query &block
		end
	end

	class DataAccessObject < AbstractAccessObject
		def where
			"[#{@table.Name}].[#{@field.Name}](Rec ##{@recno})"
		end
		
		def iterator(&block)
			each_datum &block
		end

    def replace(new_value, re)
			begin
				@rs.Edit
				@field.Value = @field.Value.to_s.gsub!(/#{re}/, new_value)
				@rs.Update
			rescue StandardError => oops
				$stderr.puts "Exception replacing [#{@table.Name}]({@recno}).[#{@field.Name}].Value = '#{@field.Value.to_s}': #{oops}"
			end
    end
		
		def delete_current_line
			begin
				@rs.Delete
			rescue StandardError => oops
				$stderr.puts "Exception deleting [#{@table.Name}]({@recno}).[#{@field.Name}]: #{oops}"
			end
		end

		def each(&block)
			each_datum do |table|
				sql = "SELECT #{Options.fields.empty? ? '*' : Options.fields.join(',')} FROM [#{table.Name}]"
				sql += " WHERE#{Options.where_clause}" unless Options.where_clause.empty?
				$stderr.puts ">>> Executing SQL \"#{sql}\"" if Options.verbose
				begin
					@rs = @access.CurrentDb.OpenRecordset(sql)
					@recno = 0
					while not @rs.EOF do
						@recno += 1
						@rs.Fields.each do |@field|
							yield @field.Value.to_s
						end
						@rs.MoveNext
					end
				rescue StandardError => oops
					$stderr.puts "Exception processing #{table.Name}: #{oops}"
					@rs.Close
					@rs = nil
				end
			end
		end
	end
	
	class ReferencesAccessObject < LineAccessObject
    def each
      @line = 0
      @access.Application.VBE.ActiveVBProject.References.each do |ref|
        @line += 1
        yield "#{ref.Name} -- #{ref.Description rescue "<no description>"} (path: #{ref.FullPath})"
      end
    end
  end

  class MacroAccessObject < LineAccessObject
    def where_line
      "#{@comp}!##{@line}"
    end
    
    def replace(line)
      comp = @access.VBE.ActiveVBProject.VBComponents(@comp)
      comp.CodeModule.ReplaceLine(@line, line)
    end

    def delete_current_line
      comp = @access.VBE.VActiveBProject.VBComponents(@comp)
      comp.CodeModule.DeleteLines(@line, 1)
    end

    def each
      @access.VBE.ActiveVBProject.VBComponents.each do |comp|
        @comp = comp.Name
        code = comp.CodeModule.Lines(1, 65536)
        @line = 0
        code.split(/\r?\n/).each do |line|
          @line += 1
          yield line
        end
      end
    end
  end

  class ProcAccessObject < MacroAccessObject
    attr_accessor :proc
    
    def where_line
      "#{@comp}.#{@proc}!##{@line}"
    end

    def each
      @access.VBE.ActiveVBProject.VBComponents.each do |comp|
        @comp = comp.Name
        begin
          @line = comp.CodeModule.ProcBodyLine(@proc, 0)
          cnt = comp.CodeModule.ProcCountLines(@proc, 0)
          code = comp.CodeModule.Lines(@line, cnt)
          code.split(/\r?\n/).each do |line|
            yield line
            @line += 1
          end
        rescue
          @line = 0
        end
      end
    end
  end

 class << self
    def new(object_type, access)
      klass = 
        case object_type.downcase
          when /^macro/
            MacroAccessObject
          when /^proc/
            ProcAccessObject
          when /^prop|^con|^ctrl/
            ControlAccessObject
          when /^addin/
            AddInAccessObject
          when /^ref/
						ReferencesAccessObject
          when /^rep/
            Options.controls.empty? ? ReportPropertyAccessObject : ReportControlAccessObject
					when /^form/
            Options.controls.empty? ? FormPropertyAccessObject : FormControlAccessObject
					when /^tab/
						TablePropertyAccessObject
					when /^quer/
						QueryPropertyAccessObject
					when /^dat/
						DataAccessObject
					end
      klass::new(access)
    end
  end
end
