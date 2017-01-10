require 'simple_xlsx_reader'

foodDocList = Dir.entries("./Food Docs")

def readInfoFromFile(fileName)
	puts fileName
	# begin
	doc = SimpleXlsxReader.open('./Food Docs/'+fileName)
	firstSheetRows = doc.sheets.first.rows
		# recipeName =  firstSheetRows[0][1]
		# puts "recipeName: ", recipeName
		# if recipeName != ''
	startParse = false;
	blankCellCount =0;
	
	firstSheetRows.each do |row|
		if row[0] =~ /store/i
			startParse=true
		end

		if startParse
			checkAndParse(row)
		end

		if startParse and (!row[0].is_a? String or row[0].strip=='')
			blankCellCount+=1
		end

		if startParse and blankCellCount>=2
			break
		end
	end
		# end
	# rescue Exception => msg
	# 	puts "error: ",msg
	# end
end

def checkAndParse(row)
	store = getStoreName(row[0]);
	partMeal = row[1]
	item = row[2]
	quantify = row[3]
	unit = row[4]
	# if !store.is_a? String or store.strip=='' or store =~ /.*(ESTIMAT|number|store|actual|date|recipe|budget|charge|fh|EVALUATION|item|NA).*/i 
	# 	return true
	# end

	# return getStoreName(store);

end

def getStoreName(store)
	case store
	when /.*costco.*/i
		store = 'Costco'
	when /.*(vege|veggi).*/i
		store = 'Vegetable'
	when /.*rt.*/i
		store = 'RT'
	when /.*(meat|pork).*/i
		store ='Pork Vendor'
	when /.*america.*/i
		store ='America'
	when /.*亦慧.*/i
		store ='亦慧'
	when /.*(v|~).*/i
		store ='undetermined'
	else
		puts "other types: #{store} appears!"
		return false;
	end
	return store;
end


foodDocList = foodDocList - ['.','..','.DS_Store'];

foodDocList.each do |fileName|
	readInfoFromFile(fileName)
end


