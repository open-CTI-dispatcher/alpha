#  This Source Code Form is subject to the terms of the Mozilla Public
#  License, v. 2.0. If a copy of the MPL was not distributed with this
#  file, You can obtain one at http://mozilla.org/MPL/2.0/.

#  This Source Code Form is "Incompatible With Secondary Licenses", as
#  defined by the Mozilla Public License, v. 2.0.



require 'sinatra'
require 'rb-scpt'
include Appscript
require 'pathname'


class MyApp < Sinatra::Application

	configure  do
		disable :sessions, :protection
    disable :logging
		set :environment, :production
		set :server, :thin
		set :bind, '10.10.10.42'
		set :port, 8080
		set :lock, true

		set :calls, {}
		set :music_playing, false
		set :CalendarID, 328
		set :categories, {
			:in  => [13, 59, "Anruf eingehend"],
			:out => [14, 60, "Anruf ausgehend"],
			:nop => [15, 61, "unbeantwortet"],
			:mis => [16, 62, "verpasst"],
			:con => [17, 63, "angenommen"],
			:end => [18, 65, "beendet"]}
			# sym    index, Outlook category-ID, text 

		### write log of all calls. ensure log-directory exists
		set :logpath, Pathname.new(ENV["HOME"]) + "Library/Logs/CTI"
		Dir.mkdir( settings.logpath ) if not Dir.exists? settings.logpath
		### write also a debug-log? 
		set :debug_log, true

		### Music.app has no rights in VS Code, 
		### so its call crashes => detect and disable
		if app_file == Pathname.new(ENV["HOME"]) + "Entwicklung/CTI/cti.sinatra.rb"
			set :environment, :development 
			$stdout.sync = true		# flush output
		end

		#debug_log("Startup CTI complete.")
  end

	# some call patterns:
	# sp: speaker - hs: handset
	#
	# in+missed:	incomingcall                                                  > callterminated > missedcall
	# in+sp:			incomingcall           > callestablished                      > callterminated 
	# in+sp+hs:		incomingcall           > callestablished > offhook > (onhook) > callterminated > (onhook)
	# in+hs:			incomingcall > offhook > callestablished >            onhook  > callterminated
	#
	# out+missed:	          outgoingcall                            > callterminated
	# out+sp:			          outgoingcall > callestablished          > callterminated
	# out+hs:			offhook > outgoingcall > callestablished > onhook > callterminated
	#
	# offhook > onhook can take place at any time (except incoming missed calls) 
	# during calls all relevant fields are set
	# onhook has no informations after "callterminated"
	#
	# http://10.10.10.42:8080/incomingcall/snom?device=$mac&activeline=$active_line&phonetype=$phone_type&displayremote=$display_remote&displaylocal=$display_local&outgoingidentity=$outgoing_identity&userrealname1=$user_realname1&userrealname2=$user_realname2&activeuser=$active_user&nr_of_calls=$nr_ongoing_calls&activeurl=$active_url&local=$local&remote=$remote&incomingidentity=$incoming_identity
	#
	# http://10.10.10.42:8080/incomingcall/snom?phonetype=$phone_type&device=$mac&activeline=$active_line&nr_of_calls=$nr_ongoing_calls&local=$local&displaylocal=$display_local&displayremote=$display_remote
	#
	#  Snom action url parameters 
	# phonetype: snomD765
	# device: 0049123456789
	# activeline: 1
	# nr_of_calls: $nr_ongoing_calls
	# local: cug_snom@10.10.10.1
	# displaylocal: Tom
	# outgoingidentity: 1
	# displayremote: iPhone **621

	# retrieve list of Outlook categories 
	# app("/Applications/Microsoft Outlook.app").categories.get.each{|a|(ae<<[a.id_.get,a.name.get]) if (a.name.get[0..4]=="Anruf")};ae
	# all Categories and ids are stored in "ae"

	# dial at snom phone via http
	# http://admin:passwort@10.10.10.104/command.htm?number=1234&outgoing_uri=cebusnom@10.10.10.4

	
	
	############################
	def get_call_data(direction)
		# record data & time
		call = {}  
		call[:dir]  =  direction
		call[:time] =  Time.now
		settings.calls[pli] = call
	end
	
	get '/outgoingcall/snom' do
		debug_log("outgoingcall")
		get_call_data("out")
		[200, {}, ""]
	end	
	
	get '/incomingcall/snom' do
		debug_log("incomingcall")
		get_call_data("in")
		[200, {}, ""]
	end	
	

	##############################
	# mark call as connected & time
	get '/callestablished/snom' do
		
		if ( not settings.calls.include?(pli) )
			debug_log(" pli leer => raus")
			return [200, {}, ""]
		end
		
		debug_log("callestablished")
		settings.calls[pli][:con] = Time.now
		stop_music
		
		# create calendar event in Outlook and show it
		subject = ""
		subject << "Anruf"
		subject << (incoming? ? " von " : " bei ")
		subject << par('displayremote')
		subject << (incoming? ? " für " : " als ")
		subject << par('displaylocal')
		text    = "§\n----------\n\n"

		create_event(subject, text, true)
		add_category( incoming? ? :in : :out )
		add_category( :con )

		[200, {}, ""]
	end
	

	#############################
	# record data & time; clear phone/line data
	get '/callterminated/snom' do

		# "declined" calls trigger callterminated() 2! times!? 
		# first time "regular", 2 seconds later once again!?
		if  ( not settings.calls.include?(pli) )
			# no pli-array in call data => already deleted
			#debug_log("2rd trigger - no index")
			return [200, {}, ""]
		end

		debug_log("callterminated/go ")
		
		call = settings.calls[pli]
		call[:end] = Time.now
		resume_music
		
		# put data of call into calendar event in Outlook
		if call.include?(:con)
			
			##########################
			# call connected => calendar_event exists => update call data
			cinfo = ""
			cinfo << call[:time].strftime("%A, %d. %B %Y um %k:%M")          << "  -  " 
			cinfo << ( incoming? ? "Anruf eingegangen von " : "Angerufen bei ")
			cinfo << par('displayremote') 
			cinfo << ( incoming? ? " für " : " als " ) << par('displaylocal')  << "  -  " 
			cinfo << "Dauer "     << time_difference(call[:con] ,call[:end]) << "  -  "
			cinfo << "Klingeln "  << time_difference(call[:time],call[:con]) << "  -  "
			cinfo << "Identität " << par('displaylocal') << " (" << par('local') << ")"
			
			debug_log("update calendarevent " + call[:ev_id].to_s)
			cev  = Appscript::app("Microsoft Outlook").calendar_events.ID(call[:ev_id])
			txt = cev.plain_text_content.get.gsub(/§/, cinfo)
			cev.plain_text_content.set(txt)
			cev.end_time.set(	call[:end] )
			add_category( :end )
			call[:type] = :con
			debug_log("update calendarevent finished")
			
		else
			##########################
			# call had no connection => no calendar event => create one
			header  = (incoming? ? "Anruf verpasst von " : 
														 "Anruf unbeantwortet an ")
			subject = header<<par('displayremote')<<(incoming? ? " für " : " als ")<<par('displaylocal')

			cinfo = ""
			cinfo << call[:time].strftime("%A, %d. %B %Y um %k:%M") << "  -  " 
			cinfo << header << "  -  " 
			cinfo << "Klingeln "  << time_difference(call[:time],call[:end]) << "  -  "
			cinfo << "Identität " << par('displaylocal') << " (" << par('local') << ")"
			cinfo << "\n----------\n\n"

			create_event(subject, cinfo, false)
			add_category( incoming? ? :in  : :out )
			add_category( incoming? ? :mis : :nop )
			call[:type]=( incoming? ? :mis : :nop )			
		end
		
		
		# call finished - log call and delete data for this phone & line
		call_log
		settings.calls.delete(pli)
		debug_log("callterminated/end")
		[200, {}, ""]
	end


	#########################
	get '/missedcall/snom' do
		# called after "callterminated"
	end

	get '/offhook/snom' do
	end

	get '/onhook/snom' do
	end

	get '/WEG/:what/snom' do
		puts "**************** what: " + params['what']
		#puts "**************** para"
		params.each{|p,v|puts p.to_s + ": " + v.to_s  }
		#puts request.path_info
		#puts request.path
		#puts request.ip
		puts "**************** ende"
		[200, {}, ""]
	end
	

	#########################
	### Stop Thin
	get '/exit' do
  	exit!
	end


	#########################
	def create_event(subject, text, open = false)
		debug_log("create event/go ")
		
		Appscript::app("Microsoft Outlook").launch
		cid = settings.CalendarID
		cal = Appscript::app("Microsoft Outlook").calendars.ID( cid ).make(
			:new=>:calendar_event,
			:with_properties=>{
				:location=> "Identität "+par('displaylocal')+" von "+par('phonetype')+" auf Leitung "+par('activeline'),
				:subject=>subject,
				:plain_text_content=>text,
				:start_time=>Time.now,
				:end_time=>Time.now,
				:has_reminder=>false} )
		settings.calls[pli][:ev_id] = cal.id_.get
		cal.open if open
		debug_log("create event/end")
		cal.id_.get
	end

	def add_category(cat_sym)
		eid  = settings.calls[pli][:ev_id]
		cev  = Appscript::app("Microsoft Outlook").calendar_events.ID(eid)
		cas  = cev.categories.get
		cas << Appscript::app("Microsoft Outlook").categories.get[ settings.categories[cat_sym][0] ]
		cev.categories.set(cas)
	end


	#########################
	def stop_music()
		if settings.environment == :production
			settings.music_playing = Appscript::app("Music").player_state.get
			Appscript::app("/System/Applications/Music.app").playpause
		end
	end
	
	def resume_music()
		if settings.environment == :production
			if settings.music_playing == :playing
				if Appscript::app("Music").player_state.get==:paused
					Appscript::app("Music").playpause 
				end
			end
		end
	end


	#########################
	def write_header(pf)
		# write csv-header 
		pf.open('a') {|f| 
			f.write("Date, Time, Direction, Status, Number, Duration, Ring Duration, Identity, Phone&Line\n")
			f.flush
		}
	end
		
	## logger
	def call_log(text = "")
		call  = settings.calls[pli]
		cinfo = ""
		cinfo << call[:time].strftime("%d. %B %Y, %k:%M") + ", "
		cinfo << ( incoming? ? "ein" : "aus" ) + ", "
		cinfo << settings.categories[ call[:type] ][2] + ", "
		cinfo << par('displayremote')  + ", "
		if call.include?(:con)
						 # connected call 
		cinfo << time_difference(call[:con] ,call[:end]) + ", "
		cinfo << time_difference(call[:time],call[:con]) + ", "
		else	
			       # no cennection => no call duration => only ring duration
		cinfo << "00:00, "
		cinfo << time_difference(call[:time],call[:end]) + ", "
		end
		cinfo << par('userrealname' + par('outgoingidentity')) + 
		" (" + par('outgoingidentity') +"), "
		cinfo << par('phonetype') + " line " + par('activeline')
		cinfo << ", " + text if (not text=="")
		cinfo << "\n"

		fn = "cti.log." + Time.now.strftime("%Y-%m-%d") + ".csv"
		pf = Pathname.new( settings.logpath ) + fn
		write_header(pf) unless File.exists?(pf)
		pf.open('a') {|f| 
			f.write(cinfo)
			f.flush
		}
	end##def call_log


	def debug_log(text = "")
		cpara=""
		params.map{|p,v| cpara << p.to_s + ": " + v.to_s + " / "  }

		puts Time.now.strftime("%k:%M:%S:%L ") + text.ljust(18," ") + " " +
					par('displaylocal').ljust(10," ") + " " +
					par('local') #+ "\n" + cpara

		return if not settings.debug_log

		cinfo  = ""
		cinfo << Time.now.strftime("%Y %m %d, %k:%M:%S:%L,")
		if (not text==""); cinfo << text + ","; end
		cinfo << cpara
		cinfo  = cinfo[0..-3] + "\n"

		fn = "cti.debug." + Time.now.strftime("%Y-%m-%d") + ".csv"
		pf = Pathname.new( settings.logpath ) + fn
		pf.open('a') {|f| 
			f.write(cinfo)
			f.flush
		}
	end##def debug_log


	############################
	def par(was)
		#prevent empty params
		params[was].nil? ? "" : params[was]
	end

	def pli() #phone-line-identity-nr
		#key in active-calls-hash of current call data
		params['device']+"-"+params['activeline']+"-"+params['local']+"-"+params['displayremote']
	end 

	def incoming?() 
		settings.calls[pli][:dir]=="in"		# in=true / out=false
	end

	def time_difference(start_t = Time.now ,end_t = Time.now)
		start_t = Time.now if start_t.nil?
		end_t   = Time.now if   end_t.nil?
		time_diff = (end_t - start_t).round.abs
		hours = time_diff / 3600
		dt = DateTime.strptime(time_diff.to_s, '%s')
		hours>0 ? "#{hours}:#{dt.strftime "%M:%S"}" : "#{dt.strftime "%M:%S"}"
	end

	#########################
	# start the server
	puts "*********** \nApp_file >" + app_file.to_s + "< \n$0:      >" + $0 + "<\n"
	begin
		run! if app_file == $0

	rescue StandardError => e
		if (e.inspect.include?("port is in use"))
			puts "\n - Port already in use - not starting -\n\n"
			$stdout.flush
			exit!
		end
	end
end
