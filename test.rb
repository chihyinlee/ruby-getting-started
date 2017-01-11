require 'simple_xlsx_reader'
require 'mongo'

client = Mongo::Client.new([ '127.0.0.1:27017' ], :database => 'FOOD_DOC')
db = client.database
collection = client[:DOCS]


foodDocList = Dir.entries("./Food Docs")

GDriveRecipeMapping = {  
   "ZhaJiang Noodles 炸醬麵.xlsx"=> "https://docs.google.com/spreadsheets/d/1JR0uwdcxc0yh9CrPSUf-h2cbU8Z7vS3ADT1dqIzj8es/edit?usp=drivesdk",
   "Waffle_sis sleepover.xlsx"=> "https://docs.google.com/spreadsheets/d/1j0ec_84iCJnunWUP5GvqyeLBYfPLMRpiCqHAvsjTr88/edit?usp=drivesdk",
   "Vietnamese_Pho.xlsx"=> "https://docs.google.com/spreadsheets/d/1r1LO6XYexQHxX5GLsd1u68ttREstXTxDvl0yLsWt-YM/edit?usp=drivesdk",
   "ThaiMincedPork_FNP 打拋豬肉.xlsx"=> "https://docs.google.com/spreadsheets/d/1TG05XSi_1hbt8ZHAwkUglj9VImgTnWw-ALczbpfkT9c/edit?usp=drivesdk",
   "Three Cup chicken 三杯雞_.xlsx"=> "https://docs.google.com/spreadsheets/d/1uSVWld1p9SIuQphk0bnvNB6huvlj6rCr9i1IUUtndSs/edit?usp=drivesdk",
   "vietnamesebibimbap_sws 越南拌飯.xlsx"=> "https://docs.google.com/spreadsheets/d/1wEsVU_95VS33LZihnRl5_a6Z3pUtfTC28Api2V4tpdA/edit?usp=drivesdk",
   "ThaiCurry 泰式咖哩.xlsx"=> "https://docs.google.com/spreadsheets/d/198He17CHX-qJ_P_6saFPGgVSRPzFHmg2rs7b5r1_Tkc/edit?usp=drivesdk",
   "Teriyaki Pork Shoulder Bake 烤豬肉肩.xlsx"=> "https://docs.google.com/spreadsheets/d/1amSeGTgUH7vgpw0yHzl2jZzxxyYIc41teTS6DVDjgxQ/edit?usp=drivesdk",
   "Sushi Bake - 壽司焗烤飯.xlsx"=> "https://docs.google.com/spreadsheets/d/11omH9CdhNjB2poo0lnoLZFST6Gx2OZTZ_TnRP9aTBmI/edit?usp=drivesdk",
   "Teriyaki Chicken.xlsx"=> "https://docs.google.com/spreadsheets/d/162g2J36IFzOQtZLP5T_B-rYDJgRwrT7m2qlXG5LpNdU/edit?usp=drivesdk",
   "SWS Snacks.xlsx"=> "https://docs.google.com/spreadsheets/d/1k6Vrstb6V8BffDAx2kdtb9gzBJiS26_ftKAayb5CbK8/edit?usp=drivesdk",
   "Sushi Bake - Christian Gathering 壽司焗烤飯.xlsx"=> "https://docs.google.com/spreadsheets/d/1h0des08ovwavQWCEyUTUxFEhkkAK0LHLdbtiH2b5vJc/edit?usp=drivesdk",
   "Stir-fried pumpkin Mifen (南瓜米粉).xlsx"=> "https://docs.google.com/spreadsheets/d/1FSJ1OvRNzfz2KeIS5jEuCSLRSdOYMDP6NOiNR0sc-9E/edit?usp=drivesdk",
   "SoonDoBu_SWS.xlsx"=> "https://docs.google.com/spreadsheets/d/16u60yUGyG-3bNvAUi-v5s21JHwLEbpkvnvRUxh4SBRw/edit?usp=drivesdk",
   "Shepherd's Pie.xlsx"=> "https://docs.google.com/spreadsheets/d/17TdBZsif5wkL8E7ll9DZwb-5wqaL4DYoi1ZBfsU_MLg/edit?usp=drivesdk",
   "Soysauce Ground Pork over Rice 滷肉飯.xlsx"=> "https://docs.google.com/spreadsheets/d/1m3rimYdjNtbp-BrMY0-MCPfQ7Tx2OKHjKCwVQbyJ1CQ/edit?usp=drivesdk",
   "ShaCha Pork Noodle 沙茶豬肉拌麵.xlsx"=> "https://docs.google.com/spreadsheets/d/1dnHlKEKYcbKxvEqvh_LoH0sVGAV_M_wW6ls73cFfGrw/edit?usp=drivesdk",
   "RudysRubChicken_SWS.xlsx"=> "https://docs.google.com/spreadsheets/d/1FbYAx8GVSn29OlnaNYN7wmNZR8BxVlO8Dwg0F6UQEv8/edit?usp=drivesdk",
   "Sausage and Chicken Pasta Alfredo.xlsx"=> "https://docs.google.com/spreadsheets/d/1Kb5FTPG8_zfpJc59uSLr6PJnQDCctwwDQu83MdPj6lg/edit?usp=drivesdk",
   "SaltWaterChicken_SWS 鹽水雞.xlsx"=> "https://docs.google.com/spreadsheets/d/1TCOWRMBICeI048-m4AAuuCGjIqNl-KRu1afs9g0jhBI/edit?usp=drivesdk",
   "Rotisserie Chicken.xlsx"=> "https://docs.google.com/spreadsheets/d/1inyto0DdB6jLbG8ltdGJ9nqnB2-USJwT2bbNocnT9LQ/edit?usp=drivesdk",
   "Roasted Vegetables.xlsx"=> "https://docs.google.com/spreadsheets/d/17UulZnXz1yPZWp9F6ZU4RVfsTV0KWL0TftzLT5uSsN4/edit?usp=drivesdk",
   "PulledPorkSandwich.xlsx"=> "https://docs.google.com/spreadsheets/d/1YQ3zmbDSTmJbki6ylAoU9LcMmKTA1IM9FybYZuatITs/edit?usp=drivesdk",
   "PorkwDaikon 燉豬肉和白蘿菠.xlsx"=> "https://docs.google.com/spreadsheets/d/1E5rf43qkp-WnjqiQfiW6KPJl33ybeQI0cOPxxfNa6Fc/edit?usp=drivesdk",
   "Post TFN snacks.xlsx"=> "https://docs.google.com/spreadsheets/d/1RhcERblcC8Ewpgt2jKX1mMEncsbCPiXF6TZ60JJjvU4/edit?usp=drivesdk",
   "porkstew 美式燉豬肉.xlsx"=> "https://docs.google.com/spreadsheets/d/1ssR6nOj2KB0ywutZDDCWVE8zHife4jt3n5lG63OivzI/edit?usp=drivesdk",
   "PIzza MYO_.xlsx"=> "https://docs.google.com/spreadsheets/d/17GDtriOSw8ZylIXp0oJ8vY6qdkoW8L8OYK4poRaLkb0/edit?usp=drivesdk",
   "Pink Sauce Pasta and salad SWS 粉紅醬義大利面.xlsx"=> "https://docs.google.com/spreadsheets/d/11tic6NoAwKyZSqI50B0pZtOqqnXPA9HrNZ1QhO6xA30/edit?usp=drivesdk",
   "PhillyCheeseSteakSandwich_FNP 費城起司潛艇堡.xlsx"=> "https://docs.google.com/spreadsheets/d/11gsttCUZBMdaAg-3BULPQcwBU0mGJf24pPKXNO4daZs/edit?usp=drivesdk",
   "PastawGroundPorkoverRice_LG 義大利蓋飯.xlsx"=> "https://docs.google.com/spreadsheets/d/1Z7aA6qEGwUzUMRpSzRWrbiXfaNJhdlExdcH4XsYHPGw/edit?usp=drivesdk",
   "Pasta with Clam 蛤蜊義大利麵.xlsx"=> "https://docs.google.com/spreadsheets/d/1fFklGX35iL4oGA2OZ7Mdu6i_Y-qnt3bhecKZz6bbhvE/edit?usp=drivesdk",
   "Pasta Salad_.xlsx"=> "https://docs.google.com/spreadsheets/d/1E2jLTPoCF-_BrM9h0zBSy4zCzEJV2yGyAiiq_Bu89S0/edit?usp=drivesdk",
   "Pasta RedSauce w_ GroundPork_HomemadeSauce.xlsx"=> "https://docs.google.com/spreadsheets/d/1VhKeFWdkBUfFuR8YO03bOzHQJi6kFW0ISOi_zKmCXa8/edit?usp=drivesdk",
   "OysterSaucePorkOverRice_蠔油豬肉飯.xlsx"=> "https://docs.google.com/spreadsheets/d/1emw5zcjtt6Zz8p620EXg7Gl03TOSiWVNkZaQNdIRypY/edit?usp=drivesdk",
   "MYO Waffles.xlsx"=> "https://docs.google.com/spreadsheets/d/1k369gS13kaVgwkRTnI4pLDMAVEtEfKhnzTw6Rk4ZjeM/edit?usp=drivesdk",
   "Pad See Ew 泰式河粉.xlsx"=> "https://docs.google.com/spreadsheets/d/1n1DpteE0PSTbo9_gvqg8hKsNQjXybIgaT_uWvpxZFjs/edit?usp=drivesdk",
   "MYO Burritos_Version2.xlsx"=> "https://docs.google.com/spreadsheets/d/1pTK1l5EDAcPGG1Mh2YoQlxGW0BNf_-BZZIqaXFvyLGs/edit?usp=drivesdk",
   "MYO Burritos.xlsx"=> "https://docs.google.com/spreadsheets/d/1ChG_nsCNdSonOUSpF4gg369VjAqVy-QPsTfHqLOU76Q/edit?usp=drivesdk",
   "MeatLoaf + Mashed Potatoes.xlsx"=> "https://docs.google.com/spreadsheets/d/1R4ThOV_qtOQFllKuxrLjB4hdbycBemHHLfDoHw6lqmE/edit?usp=drivesdk",
   "MontrealChicken.xlsx"=> "https://docs.google.com/spreadsheets/d/1ulkIJR3CP0dLtwvxSkCvNZDY57OHNw8DfyoqVM_auUI/edit?usp=drivesdk",
   "MontrealChicken_LG.xlsx"=> "https://docs.google.com/spreadsheets/d/1zZ9LrVZEoDeUf4h0ykgxTXxMWmQUkeby5OzJ4-jopeU/edit?usp=drivesdk",
   "Miso Baked Pork 味增烤豬肉.xlsx"=> "https://docs.google.com/spreadsheets/d/1GK5bC6Ya4YhYbukusbBD8XsWHH7osflWX6fW75sVdjk/edit?usp=drivesdk",
   "MapoTofu_FNP.xlsx"=> "https://docs.google.com/spreadsheets/d/1gQVYXz2mqQTKHTMtL0IEr3PrQtN-51f0fys2m4fUBAw/edit?usp=drivesdk",
   "Kung Pao chicken 宮保雞丁.xlsx"=> "https://docs.google.com/spreadsheets/d/15gVRDdzf7zqfGycgicOpuYsi_utUhxvSePLqCmhyu8Y/edit?usp=drivesdk",
   "KoreanBBQ_StaffDinner 韓式燒烤.xlsx"=> "https://docs.google.com/spreadsheets/d/1Qj9EFvTp95WR_MbjGwIvECjFNR29Nn73EENiovloIbc/edit?usp=drivesdk",
   "Korean Pork Stir-fry 韓國炒豬肉.xlsx"=> "https://docs.google.com/spreadsheets/d/1SNe0IZE5dKwT2qfOr9tVwLrv5HjFcv09SRI9PnyQk20/edit?usp=drivesdk",
   "Korean Pork w_ Tofu + seaweed banchan.xlsx"=> "https://docs.google.com/spreadsheets/d/1e7KdmGwsbZuYqHLufA59nYgXCAvagb-td0ixk9Tpm8g/edit?usp=drivesdk",
   "kimchi jigae泡菜湯.xlsx"=> "https://docs.google.com/spreadsheets/d/1Xv6riGeB_1Dq1jO1TlKUG4nxbi2wrflsAWF5jx8oxPY/edit?usp=drivesdk",
   "Hamburger BBQ.xlsx"=> "https://docs.google.com/spreadsheets/d/134NF2_ixTm-Gc2zRCBndhaTeaU5Sv6con-ltjdAv1jE/edit?usp=drivesdk",
   "JapaneseChickenCurry 日式咖哩.xlsx"=> "https://docs.google.com/spreadsheets/d/1D74hNQLpg6aCRuHjbXoXlZStolmKmfcB5BPwSqjR9xA/edit?usp=drivesdk",
   "Hummus Dip.xlsx"=> "https://docs.google.com/spreadsheets/d/1l_tHbLmd-QvBYxYaLC_goHO2oFtN46omeyh62cVnNxU/edit?usp=drivesdk",
   "Hamburger BBQ(1).xlsx"=> "https://docs.google.com/spreadsheets/d/1cxSmEQEYlOl3aSXd7XFApUJYW5Xw8Np1ZrRZNqYhUeA/edit?usp=drivesdk",
   "Grandma's Chicken Noodle Soup.xlsx"=> "https://docs.google.com/spreadsheets/d/1Aq1KNAklEtT-w24ReE_sRbhw-r55jFsznNLT-rGdPto/edit?usp=drivesdk",
   "Grandma's Corn Chowder Soup.xlsx"=> "https://docs.google.com/spreadsheets/d/1EtOJLcBnNG12p-HiqoZxyu1y1RlFp0n7VdEgTYYA3uc/edit?usp=drivesdk",
   "Gumbo 美國南部濃湯.xlsx"=> "https://docs.google.com/spreadsheets/d/1iBz6dli21aZwXi0JCL9KI5Y6S_duBt6cx_8UvSabxHE/edit?usp=drivesdk",
   "FishTacos.xlsx"=> "https://docs.google.com/spreadsheets/d/1mRsbgGDHjRJSYPsIae6wuG0xzvswY0oRsE1mx8UmmzI/edit?usp=drivesdk",
   "CuminPork_StaffDinner 孜然烤豬肉.xlsx"=> "https://docs.google.com/spreadsheets/d/1ZBfr47NjmP63FUkWPptT_UAnoXao3196J37KnEuAEm8/edit?usp=drivesdk",
   "CreoleChicken_LG.xlsx"=> "https://docs.google.com/spreadsheets/d/1a4odl4UH_NfisgHylZWz6oZCrboCXm9AXqeo1akvKNE/edit?usp=drivesdk",
   "Cranberry Chicken Salad Filling.xlsx"=> "https://docs.google.com/spreadsheets/d/1Mm8lx7hG8tWhYF4Sc9PGTB322dtYCW7fK3F0mA6Bu_E/edit?usp=drivesdk",
   "Creole Chicken.xlsx"=> "https://docs.google.com/spreadsheets/d/1U7B4CMmGkdCGXn649eQHenPmu1hqgKuAR6xp8nn60jc/edit?usp=drivesdk",
   "ChickenStirfry 雞肉炒蔬菜.xlsx"=> "https://docs.google.com/spreadsheets/d/1sZE6g6ILqRMsjyEm8OxsUwWPM3ok_EbSt7wfrHN-Tz8/edit?usp=drivesdk",
   "ChickenJjim  韓式醬油雞.xlsx"=> "https://docs.google.com/spreadsheets/d/10fww7dbFIuEW19GWASI3SlHIe46WPFHLGJwXi7o3zWE/edit?usp=drivesdk",
   "ChickenEnchiladas_Discipleship_20120414.xlsx"=> "https://docs.google.com/spreadsheets/d/153wubfehd7xrXKAn65n6jCOC3iezDt7GIE-PCObtkbs/edit?usp=drivesdk",
   "Chicken Soba Noodle Salad_.xlsx"=> "https://docs.google.com/spreadsheets/d/17u7anFFB0rPDBx4DXYPcViDLcduZRH5xnneUo1zRu-Y/edit?usp=drivesdk",
   "Chicken Stew.xlsx"=> "https://docs.google.com/spreadsheets/d/1A7TeihkWBaYsZR72Bsl_BJ-0uqT_owFjGthBoQBqozI/edit?usp=drivesdk",
   "ChickenGumbo 20131028 美國南部濃湯.xlsx"=> "https://docs.google.com/spreadsheets/d/1QCB8unvez90Iw5LJvQzMGcWg_oHRFDZ10hcQVFHl1Pc/edit?usp=drivesdk",
   "Chicken Sandwich with Basil Spread_.xlsx"=> "https://docs.google.com/spreadsheets/d/1_nHb8a8Oror9rPt6f-1QhPDJxVXtTasoLNZq5bXxWKc/edit?usp=drivesdk",
   "Chicken Salad Sandwich for Picnic.xlsx"=> "https://docs.google.com/spreadsheets/d/1NdEu_YLO1eXlvv0qv7B743dWV2D2j4_3fWraRC6Z5Q4/edit?usp=drivesdk",
   "Chicken Noodle Soup (American style).xlsx"=> "https://docs.google.com/spreadsheets/d/1CKgHbDMo36E44iNsa84gbVxpa6o8dUmHkMjH4Ggr5ps/edit?usp=drivesdk",
   "CampBLUE trip.xlsx"=> "https://docs.google.com/spreadsheets/d/1SyXsnIssJ7_eotouYhADfR3k-pRLoebdV6NuOCgCDBU/edit?usp=drivesdk",
   "Chicken Olive Oil Pasta.xlsx"=> "https://docs.google.com/spreadsheets/d/1XkBFTTXICsmCPOmxeNCeWE_-AMxgcYwHCGQ3W3I3dSY/edit?usp=drivesdk",
   "CampBLUE trip(1).xlsx"=> "https://docs.google.com/spreadsheets/d/1Io9c9kDfP2iEFHMKPU59Wl93mGcaHtPioZqgLKtLftM/edit?usp=drivesdk",
   "Burgers _ Fries_.xlsx"=> "https://docs.google.com/spreadsheets/d/1fCyup-dlOuNuiTXqvhYzPE3vURn7iq9QyugexqxrLvw/edit?usp=drivesdk",
   "Braised Ground Pork with Egg Seaweed 肉燥一鍋滷.xlsx"=> "https://docs.google.com/spreadsheets/d/1Hp0HTL9ZQrkFj4Uv2C69mMnEruyNjEszkZ2GCqxcJGw/edit?usp=drivesdk",
   "Broccoli Salad.xlsx"=> "https://docs.google.com/spreadsheets/d/1fMhl2FVxYl8pASXzY5Z4bmvpKNX0y6w7_Wx8mJh0htI/edit?usp=drivesdk",
   "BulgogiSandwich_FNP 沙威瑪.xlsx"=> "https://docs.google.com/spreadsheets/d/1uwrIdV5Rqp3Ffx03iF1S3UpZvh79KD-sBvCbAiKLB8I/edit?usp=drivesdk",
   "BorschSoup_StudyHall_20120614.xlsx"=> "https://docs.google.com/spreadsheets/d/1Wjyj5ravlrL-Ebjxj0KmfVG2q-13n-n8-hM2cFj474w/edit?usp=drivesdk",
   "bibimbap_FNP_20120413.xlsx"=> "https://docs.google.com/spreadsheets/d/14U0Qo4hnph1fjJBYSuTGMwfnM6xE1-uUPRCglHAPIyY/edit?usp=drivesdk",
   "bibimbap_FNP 韓式拌飯.xlsx"=> "https://docs.google.com/spreadsheets/d/1EDGh6-RVFv0L6G0I1KnBPudtipJFToyctC9tmfoW0s8/edit?usp=drivesdk",
   "BakedPorkChopRice_inprogress 港式焗烤飯.xlsx"=> "https://docs.google.com/spreadsheets/d/1Z8JC7vf716wFowxRT_LDLTw953VF6fq-xv81MEL1dMk/edit?usp=drivesdk",
   "Beef Noodle Soup and Lettuce Wrap_牛肉麵＆生菜豬肉鬆.xlsx"=> "https://docs.google.com/spreadsheets/d/1vT6udGKuqX3BO2IX0IrnwTnzfjpovHfjYGZR-lt8W80/edit?usp=drivesdk",
   "Balsamic Roasted Pork Shoulder 香醋烤豬肩肉.xlsx"=> "https://docs.google.com/spreadsheets/d/145Z90PkOhPTQEgbU2EjjgUSEQH1vWJrM-FQ3aAdudcQ/edit?usp=drivesdk",
   "BakedBurger_StaffDinner.xlsx"=> "https://docs.google.com/spreadsheets/d/11Zsqa-S1oChN-ujZGR5fXNvNFTxlyIY_YE2aUE2TheI/edit?usp=drivesdk",
   "Baked Pork Chops with Barbecue Sauce.xlsx"=> "https://docs.google.com/spreadsheets/d/1607McYNZXUYd_kLFfUpoqtnyhpaFFb9cyy05KdeHa-I/edit?usp=drivesdk",
   "Baked Orange Spareribs橙汁排骨_.xlsx"=> "https://docs.google.com/spreadsheets/d/1O5kpsEAEt6fOga15hH6pURR3DtjpbvvBsbEtS3n1Q-U/edit?usp=drivesdk",
   " Tomato Cream Sauce w_ GroundPork_HomemadeSauce.xlsx"=> "https://docs.google.com/spreadsheets/d/1kq8s2JsxE0zTdTpPw2sLYu-oHIpLA1hiUcqKPpL-MKI/edit?usp=drivesdk"
}



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
	ingredientList = []
	estimation =0
	estFound=false
	
	firstSheetRows.each do |row|
		if !startParse && !estFound
			if findEstPpl(row)
				estimation = findEstPpl(row)
				estFound = true
			end
		end

		if startParse 
			parseObj = checkAndParse(row)
			if parseObj!= nil 
				ingredientList.push(parseObj)
			end
		end

		if startParse and (!row[0].is_a? String or row[0].strip=='')
			blankCellCount+=1
		end

		if startParse and blankCellCount>=2
			break
		end
		if row[0] =~ /(store|MEAT)/i
			startParse=true
		end
	end

	recipe = {
		original_file_name: fileName,
		recipe_name: fileName.gsub('_',' ').gsub('.xlsx',''),
		estimation: estimation,
		ingredients: ingredientList,
		GDriveUrl: GDriveRecipeMapping[fileName.lstrip]
	}

	# puts fileName.lstrip,'url: ',GDriveRecipeMapping[fileName.lstrip]

	return recipe

	
	# end
	# rescue Exception => msg
	# 	puts "error: ",msg
	# end
end

def findEstPpl(row)
	isFound=false
	row[0..6].each do |col|
		if col.is_a? String  and (/.*es.*ple.*/i =~ col or /.*serv.*/i =~col)
			# puts 'row: ', row.inspect
			isFound=true
			break
		end
	end
	if isFound
		row[0..6].each do |col|
			colStr = col.to_s
			if /.*[0-9].*/ =~ colStr
				# puts 'ppl: ',col.scan(/\d+/)
				return colStr.scan(/\d+/).first 
			end
		end
	end
	return false;
	# puts 'Estimation not found!'
end


def checkAndParse(row)
	store = getStoreName(row[0])
	if store || row[1] != nil 
		return  {
		store: store,
		partMeal:row[1],
		item:row[2],
		quantify: row[3],
		unit: row[4]
		}
	end

	return nil
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
	when /.*(america|FH).*/i
		store = store
	when /.*(亦慧|奕慧).*/i
		store ='亦慧'
	# when /.*FH.*/i
	# 	store='FH'
	when /.*(v|~).*/i
		store ='undetermined'
	else
		# puts "other types: #{store} appears!"
		return false;
	end
	return store;
end


foodDocList = foodDocList - ['.','..','.DS_Store'];

recipeList=[]
foodDocList.each do |fileName|
	recipe = readInfoFromFile(fileName)
	recipeList.push(recipe)
end


collection.insert_many(recipeList)

