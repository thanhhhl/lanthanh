global unmetneed "/Users/kanyaanindya/Documents/08. Prof Nawi/01. Unmet need/02. Graph"

************
**# Table A
************

import excel "$unmetneed/Template for KOBE report_05Sep21 (for graph).xlsx", ///
	first sheet(Table A) cellrange(A2:J318) clear

drop if country==""
	
*Take a portion of string variable
	*Year
		gen year1 = substr(year, 1, strpos(year,"/") - 1) 
		replace year1 = year if missing(year1)
		destring year1, replace
		
	*% Didn't/Unmet/Met need
		foreach x in noneed met unmet {
		gen `x'1 = substr(`x', 1, strpos(`x',"(") - 1) 
		replace `x'1 = subinstr(`x'1, " ", "", .)
		}
		replace noneed1 = noneed if noneed=="NA"
		destring noneed1 met1 unmet1, replace
		
*Identify duplicate
	order country survey year* weighted*
	sort country year1 weighted
	by country year1 weighted:  gen dup = cond(_N==1,0,_n)
	br if dup>0

	*select weighted and unduplicate
	gen include = 1 if weighted=="weighted" & dup==0
		*If a country has two surveys, i.e., cmwf & wvs, exclude the wvs
			replace include = 1 if weighted=="weighted" & dup>0 & survey!="WVS"
		*If a country only has unweighted estimate
			by country year1:  gen dup_country = cond(_N==1,0,_n) if weighted=="unweighted"
			replace include = 1 if dup_country==0
		
	*remove if % unmet need is missing	
		keep if include == 1 & unmet1 !=.
		

*Edit naming
	rename country country_str
	encode country_str, gen (country)	
	drop noneed unmet met year
	rename (noneed1 unmet1 met1 year1) (noneed unmet met year)
	drop dup*
	
*Year-highest % unmet need
	 egen max_unmet = max(unmet), by(country)
	 sort country max_unmet
	 by country max_unmet: gen max_year = year[_N] 
	 egen country_max = concat(country_str max_year), punct(" ") 
	
*Export to excel
	export excel using "$unmetneed/dotplot", firstrow(variables) sheet(TableA, replace) 

************
**# Table B
************

import excel "$unmetneed/Template for KOBE report_05Sep21 (for graph).xlsx", ///
	first sheet(Table B) cellrange(A2:P308) clear
	
drop if country==""
	
*Take a portion of string variable
	*Year
		gen year1 = substr(year, 1, strpos(year,"/") - 1) 
		replace year1 = year if missing(year1)
		destring year1, replace
		
	*% Unmet aged30+
		foreach x in unmet30 {
		gen `x'1 = substr(`x', 1, strpos(`x',"(") - 1) 
		replace `x'1 = subinstr(`x'1, " ", "", .)
		}
		
		drop if unmet301==""
		
	
	destring men women age30_49 age50_59 age60_69 age70 urban rural, replace
	
	*replace NA with blank
		foreach x in age30_49 age50_59 urban rural {
			replace `x' = "" if `x'=="NA" | `x'=="N/A" 
		}
		
	destring men women age30_49 age50_59 age60_69 age70 urban rural unmet301 year1, replace
		
*Identify duplicate
	order country survey year* weighted*
	sort country year1 weighted
	by country year1 weighted:  gen dup = cond(_N==1,0,_n)
	br if dup>0

	*select weighted and unduplicate
	gen include = 1 if weighted=="weighted" & dup==0
		*If a country has two surveys, i.e., cmwf & wvs, exclude the wvs
			replace include = 1 if weighted=="weighted" & dup>0 & survey!="WVS"
		*If a country only has unweighted estimate
			by country year1:  gen dup_country = cond(_N==1,0,_n) if weighted=="unweighted"
			replace include = 1 if dup_country==0
			
	*remove if % unmet need is missing	
		keep if include == 1 & unmet301 !=.
		
*Edit naming
	rename country country_str
	encode country_str, gen (country)	
	drop unmet30 year
	rename (unmet301 year1) (unmet30 year)
	drop dup*
	
*Year-highest % unmet need
	 egen max_unmet30 = max(unmet30), by(country)
	 sort country max_unmet30
	 by country max_unmet30: gen max_year = year[_N] 
	 egen country_max30 = concat(country_str max_year), punct(" ") 
	
	
*Export to excel
	export excel using "$unmetneed/dotplot", firstrow(variables) sheet(TableB, replace) 
	
	
*select only last year and age groups
	preserve
		keep if year==max_year
		rename (age30_49 age50_59 age60_69 age70) (age1 age2 age3 age4)
		reshape long age, i(country year) j(unmetage) 
		la def unmetage_  1"30–49" 2"50–59" 3"60–69" 4"70+"
		la val unmetage unmetage_
		tostring unmetage, gen(unmetage_str)
		rename age unmetneed
		drop n_unmet30 men women urban rural include
	
	*Export to excel
		export excel using "$unmetneed/dotplot", firstrow(variables) sheet(Agegroup, replace) 
	restore
	
	
	