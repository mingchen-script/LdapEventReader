# LdapEventReader
	# Read me
	# This script will convert LDAP events 1644 into Excel pivot tables for workload analysis by:
	#    1. Scan all evtx files in script directory for event 1644, and export to CSV.
	#    2. Calls into Excel to import resulting CSV, create pivot tables for common ldap search analysis scenarios. 
	# Script requires Excel 2013 installed. 64bits Excel will allow generation of larger worksheet.
	#
	# To use the script:
	#  1. Convert pre-2008 evt to evtx using later OS. (Please note, pre-2008 does not contain all 16 data fields. So some pivot tables might not display correctly.)

	# LdapEventReader.ps1 v2.11 11/29(added $g_maxThreads jobs)
	#		Steps: 
	#   	1. Copy Directory Service EVTX from target DC(s) to same directory as this script.
	#     		Tip: When copying Directory Service EVTX, filter on event 1644 to reduce EVTX size for quicker transfer. 
	#					Note: Script will process all *.EVTX in script directory when run.
	#   	2. Run script
