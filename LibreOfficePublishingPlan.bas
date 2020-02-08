'*****  VBA Macro which makes a publishing plan from a spreadsheet with article data  *****

'Data structure tracks how many articles have certain keywords, and what is the combined character amount of articles

Public Type KeywordComboCharacteristic
	KeywordCombo As String
	ArticleAmount As Integer
	CharacterAmount As Long
End Type

Sub Main
    'This procedure makes sheet with keywords, number of articles and number of characters
	'MakeKeywordsSheet
	'This procedure makes a 'publishing plan', that is, sorts keywords according to how many characters have combined articles
	'this call is redundant, because included in MakePublishingPlanSheetWithDetails
	'MakePublishingPlanSheet
	'This procedure lists articles which include the plan in question. 
	MakePublishingPlanSheetWithDetails 
	
End Sub

'Creates both summary of publishing plan to Sheet3, and
'detailed plan in Sheet4. 
'This procedure lists articles which include the plan in question. 

Sub MakePublishingPlanSheetWithDetails 

	Dim Doc As Object
	Dim Sheet1 As Object
	Dim Sheet3 As Object
	Dim Sheet4 As Object
	
	Dim Sheet1FirstColumntCell As Object
	Dim Sheet1ThirdColumnCell As Object
	Dim Sheet1FourthColumnCell As Object
	
	'Stores keyword combination = topic of possible book
	Dim Sheet3FirstColumnCell As Object
    'Stores amount of articles of possible book
    Dim Sheet3SecondColumnCell As Object
    'Stores character amount of possible book 
    Dim Sheet3ThirdColumnCell As Object
    
    Dim sKeywords() As String
	Dim sCellValue() As String ' support variable
	Dim element as Variant
	Dim element2 as Variant
	Dim element3 as Variant
	Dim iSheet1Row as Integer
    Dim iSheet3Row as Integer
    Dim iSheet4Row as Integer
    Dim iKeywordComboNumber as Integer
    Dim iArticleAmountCounter as Integer
    Dim lCharacterAmountCounter as Long
    
    Dim sMaxUnhandledKeyword as String
    Dim iMaxUnhandledArticleAmount as Integer
    Dim lMaxUnhandledCharacterAmount as Long
    
    Doc = ThisComponent
	Sheet1 = Doc.Sheets(0)
	Sheet3 = Doc.Sheets(2)
	Sheet4 = Doc.Sheets(3)
	
	sKeywords = GetKeywords()
	
	Dim HandledKeywordCombos() As String

	'This cycle counts size of the first sheet. 
	iSheet1Row = 1
	Sheet1FirstColumnCell = Sheet1.getCellByPosition(0,iSheet1Row)	    
	Do Until strcomp(Sheet1.getCellByPosition(0,iSheet1Row).String, "End") = 0 OR IsEmpty(Sheet1.getCellByPosition(0,iSheet1Row))
	    iSheet1Row = iSheet1Row + 1
	Loop

	Dim FirstSheetData(iSheet1Row) As KeywordComboCharacteristic
	
	'read first sheet data to speed up function
	iSheet1Row = 1
    For Each element in FirstSheetData
	    element.KeywordCombo = Sheet1.getCellByPosition(3,iSheet1Row).String
	    element.CharacterAmount = Sheet1.getCellByPosition(2,iSheet1Row).Value
	    iSheet1Row = iSheet1Row + 1
	Next
	
	iSheet1Row = 1
    iArticleAmountCounter = 0
    lCharacterAmountCounter = 0
    
    iSheet3Row = 1
	Sheet3FirstColumnCell = Sheet3.getCellByPosition(0,iSheet3Row)
    Sheet3SecondColumnCell = Sheet3.getCellByPosition(1,iSheet3Row)
	Sheet3ThirdColumnCell = Sheet3.getCellByPosition(2,iSheet3Row)
    
    iSheet4Row = 1
    
   'This cycle counts size of the initial data of the third sheet
	Do Until strcomp(Sheet3FirstColumnCell.String, "") = 0 OR IsEmpty(Sheet3.getCellByPosition(0,iSheet3Row))
	    iSheet3Row = iSheet3Row + 1
	    Sheet3FirstColumnCell = Sheet3.getCellByPosition(0,iSheet3Row)
	Loop
 
	Dim InitialKeywordCombos(iSheet3Row-2) As KeywordComboCharacteristic
	'Dim HandledKeywordCombos(isheet3Row-1) As String
	'KeywordComboCharacteristics store space for both initial keywords and others, index from 0. 
	Dim KeywordComboCharacteristics(iSheet3Row + Ubound(sKeywords) -1) As KeywordComboCharacteristic
	
	'read initial Keyword combos speed up function
	iSheet3Row = 1
    For Each element in InitialKeywordCombos
	    element.KeywordCombo = Sheet3.getCellByPosition(0,iSheet3Row).String
	    element.CharacterAmount = Sheet3.getCellByPosition(1,iSheet3Row).Value
	    iSheet3Row = iSheet3Row + 1
	Next  
    
    iSheet3Row = 1   
    iKeywordComboNumber = 1
    
   'This cycle initializes HandledKeywordComboCharacteristics and counts character amount of initial keyword articles
   'and storest them to Sheet3 and KeywordCombocharacteristics
	
	For each element in InitialKeywordCombos
	     If Not ArticleHasHandledKeywordCombo(element.KeywordCombo,HandledKeywordCombos) Then
			For each element2 in FirstSheetData
				If (Not ArticleHasHandledKeywordCombo(element2.KeywordCombo,HandledKeywordCombos)) AND _
				 SubWordList(element.KeywordCombo,element2.KeywordCombo) Then
	           		  'count number of articles
	   		   			iArticleAmountCounter = iArticleAmountCounter+1
							'count number of characters
						lCharacterAmountCounter = lCharacterAmountCounter + element2.CharacterAmount
					
				End If
			Next
			
			KeywordComboCharacteristics(iKeywordComboNumber).KeywordCombo = element.KeywordCombo
			KeywordComboCharacteristics(iKeywordComboNumber).ArticleAmount = iArticleAmountCounter		
	        KeywordComboCharacteristics(iKeywordComboNumber).CharacterAmount = lCharacterAmountCounter
	   		
		   			
	   	    Sheet3SecondColumnCell.value = iArticleAmountCounter		
	        Sheet3ThirdColumnCell.value =  lCharacterAmountCounter
	    	
	        'This subcycle lists articles on initial keyword topics for 4th sheet
	        iSheet4Row = iSheet4Row + 1
	    	Sheet4.getCellByPosition(0,iSheet4Row).String = element.KeywordCombo
			Sheet4.getCellByPosition(3,iSheet4Row).Value =  lCharacterAmountCounter
	    
	    	iSheet1Row = 1
	    	'Sheet1FourthColumnCell = Sheet1.getCellByPosition(3,iSheet1Row)
		
			Do Until strcomp(Sheet1.getCellByPosition(0,iSheet1Row).String, "End") = 0 OR IsEmpty(Sheet1.getCellByPosition(0,iSheet1Row))
	    		If Not ArticleHasHandledKeywordCombo(Sheet1.getCellByPosition(3,iSheet1Row).String,HandledKeywordCombos) _
	    		AND SubWordList(element.KeywordCombo,Sheet1.getCellByPosition(3,iSheet1Row).String) Then
	        		iSheet4Row = iSheet4Row + 1
	        		'Copy article name
	        		Sheet4.getCellByPosition(1,iSheet4Row).String = Sheet1.getCellByPosition(0,iSheet1Row).String
	        		'Copy article date
	        		Sheet4.getCellByPosition(2,iSheet4Row).String = Sheet1.getCellByPosition(1,iSheet1Row).String	
                	'Copy article character amount 
                	Sheet4.getCellByPosition(3,iSheet4Row).Value = Sheet1.getCellByPosition(2,iSheet1Row).Value	
                 	'Copy article keywords
                	Sheet4.getCellByPosition(4,iSheet4Row).String = Sheet1.getCellByPosition(3,iSheet1Row).String	
                	 'Copy amount of article readers
                	Sheet4.getCellByPosition(5,iSheet4Row).Value = Sheet1.getCellByPosition(4,iSheet1Row).Value	
	 
	    		End If
	    		iSheet1Row = iSheet1Row + 1
			Loop
	   	
	   		iSheet1Row = 1
			
	    	
	    	Redim Preserve HandledKeywordCombos(Ubound(HandledKeywordCombos)+1)
		    HandledKeywordCombos(Ubound(HandledKeywordCombos)) = element.KeywordCombo

	    		
	    	iSheet3Row = iSheet3Row + 1
			Sheet3SecondColumnCell = Sheet3.getCellByPosition(1,iSheet3Row)
			Sheet3ThirdColumnCell = Sheet3.getCellByPosition(2,iSheet3Row)
			
					
	    iKeywordComboNumber = iKeywordComboNumber + 1
	    iArticleAmountCounter = 0
    	lCharacterAmountCounter = 0
    	  
		End If
		
	Next element	
	
	
	'in next cycle may skip keywords handled in previous cycle
	iKeyWordStartNumber = iKeywordComboNumber
    
    Sheet3FirstColumnCell = Sheet3.getCellByPosition(0,iSheet3Row)
    
	lMaxUnhandledCharacterAmount = 1
    
    'the main cycle which makes rest of the plan. 
    'Cycle goes on until there are no more unincluded texts, that is no unsorter characters remain. . 
    
    Do Until lMaxUnhandledCharacterAmount = 0
    	'element is element in general keywordlist. This cycle defines which remaining keyword has most articles
    	'and stores that keyword with its characteristics. 
    	lMaxUnhandledCharacterAmount = 0
    	For each element in sKeywords
    	 	If len(Cstr(element)) > 0 And Not ArticleHasHandledKeywordCombo(element,HandledKeywordCombos) Then
				KeywordComboCharacteristics(iKeywordComboNumber).KeywordCombo = element
				For each element3 in FirstSheetData
					If Not ArticleHasHandledKeywordCombo(element3.KeywordCombo,HandledKeywordCombos) Then 
	            		sCellValue = Split(element3.KeywordCombo,",")
	            	
	            		'element2 is a keyword in list of keywords of a certain cell 		
	    				For Each element2 in sCellvalue
	    					If strcomp(element,Trim(element2)) = 0 Then
	    	 	 				'count number of articles
	    		   				iArticleAmountCounter = iArticleAmountCounter+1
								'count number of characters
								lCharacterAmountCounter = lCharacterAmountCounter + element3.CharacterAmount
								Exit For
							End If
						Next
					End If
				Next
			    KeywordComboCharacteristics(iKeywordComboNumber).ArticleAmount = iArticleAmountCounter		
	    		KeywordComboCharacteristics(iKeywordComboNumber).CharacterAmount = lCharacterAmountCounter
	        
	        	If lCharacterAmountCounter > lMaxUnhandledCharacterAmount Then 
	        		sMaxUnhandledKeyword =  element
	        		iMaxUnhandledArticleAmount = iArticleAmountCounter
	        		lMaxUnhandledCharacterAmount = lCharacterAmountCounter
	        	End If
	        
	    		iKeywordComboNumber = iKeywordComboNumber + 1
	            iArticleAmountCounter = 0
    			lCharacterAmountCounter = 0
    	  
			End If
		
		Next element
		
		Sheet3FirstColumnCell.string = sMaxUnhandledKeyword
		Sheet3SecondColumnCell.value = iMaxUnhandledArticleAmount		
	    Sheet3ThirdColumnCell.value =  lMaxUnhandledCharacterAmount

	    'This subcycle lists articles on remaining topics for 4th sheet
	    iSheet4Row = iSheet4Row + 1
	    Sheet4.getCellByPosition(0,iSheet4Row).String = sMaxUnhandledKeyword
		Sheet4.getCellByPosition(3,iSheet4Row).Value =  lMaxUnhandledCharacterAmount
	    
	    iSheet1Row = 1
		
		Do Until strcomp(Sheet1.getCellByPosition(0,iSheet1Row).String, "End") = 0 OR IsEmpty(Sheet1.getCellByPosition(0,iSheet1Row))
	    	If Not ArticleHasHandledKeywordCombo(Sheet1.getCellByPosition(3,iSheet1Row).String,HandledKeywordCombos) _
	    	AND SubWordList(sMaxUnhandledKeyword,Sheet1.getCellByPosition(3,iSheet1Row).String) Then
	        	iSheet4Row = iSheet4Row + 1
	        	'Copy article name
	        	Sheet4.getCellByPosition(1,iSheet4Row).String = Sheet1.getCellByPosition(0,iSheet1Row).String
	        	'Copy article date
	        	Sheet4.getCellByPosition(2,iSheet4Row).String = Sheet1.getCellByPosition(1,iSheet1Row).String	
                'Copy article character amount 
                Sheet4.getCellByPosition(3,iSheet4Row).Value = Sheet1.getCellByPosition(2,iSheet1Row).Value	
                 'Copy article keywords
                Sheet4.getCellByPosition(4,iSheet4Row).String = Sheet1.getCellByPosition(3,iSheet1Row).String	
                 'Copy amount of article readers
                Sheet4.getCellByPosition(5,iSheet4Row).Value = Sheet1.getCellByPosition(4,iSheet1Row).Value	
 
	    	End If
	    	iSheet1Row = iSheet1Row + 1
		Loop
  		
  		Redim Preserve HandledKeywordCombos(Ubound(HandledKeywordCombos)+1)
		HandledKeywordCombos(Ubound(HandledKeywordCombos)) = sMaxUnhandledKeyword
  		
	    iSheet3Row = iSheet3Row + 1
	    Sheet3FirstColumnCell = Sheet3.getCellByPosition(0,iSheet3Row)
    	Sheet3SecondColumnCell = Sheet3.getCellByPosition(1,iSheet3Row)
		Sheet3ThirdColumnCell = Sheet3.getCellByPosition(2,iSheet3Row)
		
		'may skip keywords handled in previous cycle
		iKeywordComboNumber = iKeyWordStartNumber
		iArticleAmountCounter = 0
		lCharacterAmountCounter = 0
		   
	Loop
	iSheet4Row = iSheet4Row + 1
	Sheet4.getCellByPosition(0,iSheet4Row).String = "End"
	
End Sub

	'This procedure makes a 'publishing plan', that is, sorts keywords according to how many characters have combined articles
	'this call is redundant, because included in MakePublishingPlanSheetWithDetails

Sub MakePublishingPlanSheet 
	Dim Doc As Object
	Dim Sheet1 As Object
	
	Dim Sheet1FirstColumntCell As Object
	Dim Sheet1ThirdColumnCell As Object
	Dim Sheet1FourthColumnCell As Object
	
	Dim Sheet3 As Object
	Dim Sheet3FirstColumnCell As Object
	
	Dim sKeywords() As String
	Dim sCellValue() As String ' support variable
	Dim element as Variant
	Dim element2 as Variant
	Dim element3 as Variant
	Dim iSheet1Row as Integer
    Dim iKeywordComboNumber as Integer
    Dim iArticleAmountCounter as Integer
    Dim lCharacterAmountCounter as Long
    
    Dim sMaxUnhandledKeyword as String
    Dim iMaxUnhandledArticleAmount as Integer
    Dim lMaxUnhandledCharacterAmount as Long
    
    Doc = ThisComponent
	Sheet1 = Doc.Sheets(0)
	Sheet3 = Doc.Sheets(2)
	
	sKeywords = GetKeywords()
	
	Dim HandledKeywordCombos() As String

	'This cycle counts size of the first sheet. 
	iSheet1Row = 1
	Sheet1FirstColumnCell = Sheet1.getCellByPosition(0,iSheet1Row)	    
	Do Until strcomp(Sheet1.getCellByPosition(0,iSheet1Row).String, "End") = 0 OR IsEmpty(Sheet1.getCellByPosition(0,iSheet1Row))
	    iSheet1Row = iSheet1Row + 1
	Loop

	Dim FirstSheetData(iSheet1Row) As KeywordComboCharacteristic
	
	'read first sheet data to speed up function
	iSheet1Row = 1
    For Each element in FirstSheetData
	    element.KeywordCombo = Sheet1.getCellByPosition(3,iSheet1Row).String
	    element.CharacterAmount = Sheet1.getCellByPosition(2,iSheet1Row).Value
	    iSheet1Row = iSheet1Row + 1
	Next
	
	iSheet1Row = 1
    iArticleAmountCounter = 0
    lCharacterAmountCounter = 0
    
    iSheet3Row = 1
	Sheet3FirstColumnCell = Sheet3.getCellByPosition(0,iSheet3Row)
    Sheet3SecondColumnCell = Sheet3.getCellByPosition(1,iSheet3Row)
	Sheet3ThirdColumnCell = Sheet3.getCellByPosition(2,iSheet3Row)
    
   'This cycle counts size of the initial data of the third sheet
	Do Until strcomp(Sheet3FirstColumnCell.String, "") = 0 OR IsEmpty(Sheet3.getCellByPosition(0,iSheet3Row))
	    iSheet3Row = iSheet3Row + 1
	    Sheet3FirstColumnCell = Sheet3.getCellByPosition(0,iSheet3Row)
	Loop
 
	Dim InitialKeywordCombos(iSheet3Row-2) As KeywordComboCharacteristic
	'KeywordComboCharacteristics store space for both initial keywords and others, index from 0. 
	Dim KeywordComboCharacteristics(iSheet3Row + Ubound(sKeywords) -1) As KeywordComboCharacteristic
	
	'read initial Keyword combos speed up function
	iSheet3Row = 1
    For Each element in InitialKeywordCombos
	    element.KeywordCombo = Sheet3.getCellByPosition(0,iSheet3Row).String
	    element.CharacterAmount = Sheet3.getCellByPosition(1,iSheet3Row).Value
	    iSheet3Row = iSheet3Row + 1
	Next  
    
    iSheet3Row = 1   
    iKeywordComboNumber = 1
    
   'This cycle initializes HandledKeywordComboCharacteristics and counts character amount of initial keyword articles
   'and storest them to Sheet3 and KeywordCombocharacteristics
	
	For each element in InitialKeywordCombos
	     If Not ArticleHasHandledKeywordCombo(element.KeywordCombo,HandledKeywordCombos) Then
			For each element2 in FirstSheetData
				If (Not ArticleHasHandledKeywordCombo(element2.KeywordCombo,HandledKeywordCombos)) AND _
				 SubWordList(element.KeywordCombo,element2.KeywordCombo) Then
	           		  'count number of articles
	   		   			iArticleAmountCounter = iArticleAmountCounter+1
							'count number of characters
						lCharacterAmountCounter = lCharacterAmountCounter + element2.CharacterAmount
					
				End If
			Next
			
			KeywordComboCharacteristics(iKeywordComboNumber).KeywordCombo = element.KeywordCombo
			KeywordComboCharacteristics(iKeywordComboNumber).ArticleAmount = iArticleAmountCounter		
	        KeywordComboCharacteristics(iKeywordComboNumber).CharacterAmount = lCharacterAmountCounter
	   		
		    Redim Preserve HandledKeywordCombos(Ubound(HandledKeywordCombos)+1)
		    HandledKeywordCombos(Ubound(HandledKeywordCombos)) = element.KeywordCombo
		   			
	   	    Sheet3SecondColumnCell.value = iArticleAmountCounter		
	        Sheet3ThirdColumnCell.value =  lCharacterAmountCounter
	    		
	    	iSheet3Row = iSheet3Row + 1
			Sheet3SecondColumnCell = Sheet3.getCellByPosition(1,iSheet3Row)
			Sheet3ThirdColumnCell = Sheet3.getCellByPosition(2,iSheet3Row)
			   				
	    	iKeywordComboNumber = iKeywordComboNumber + 1
	        iArticleAmountCounter = 0
    	    lCharacterAmountCounter = 0
    	  
		End If
		
	Next element	
	
	
	'in next cycle may skip keywords handled in previous cycle
	iKeyWordStartNumber = iKeywordComboNumber
    
    Sheet3FirstColumnCell = Sheet3.getCellByPosition(0,iSheet3Row)
    
	lMaxUnhandledCharacterAmount = 1
    
    'the main cycle which makes rest of the plan. 
    'Cycle goes on until there are no more unincluded texts, that is no unsorter characters remain. . 
    
    Do Until lMaxUnhandledCharacterAmount = 0
    	'element is element in general keywordlist
    	lMaxUnhandledCharacterAmount = 0
    	For each element in sKeywords
        	If len(Cstr(element)) > 0 And Not ArticleHasHandledKeywordCombo(element,HandledKeywordCombos) Then
				KeywordComboCharacteristics(iKeywordComboNumber).KeywordCombo = element
		
				For each element3 in FirstSheetData
					If Not ArticleHasHandledKeywordCombo(element3.KeywordCombo,HandledKeywordCombos) Then 
	            		sCellValue = Split(element3.KeywordCombo,",")
	            	
	            		'element2 is a keyword in list of keywords of a certain cell 		
	    				For Each element2 in sCellvalue
	    					If strcomp(element,Trim(element2)) = 0 Then
	    	 	 				'count number of articles
	    		   				iArticleAmountCounter = iArticleAmountCounter+1
								'count number of characters
								lCharacterAmountCounter = lCharacterAmountCounter + element3.CharacterAmount
								Exit For
							End If
						Next
					End If
				Next
			    KeywordComboCharacteristics(iKeywordComboNumber).ArticleAmount = iArticleAmountCounter		
	    		KeywordComboCharacteristics(iKeywordComboNumber).CharacterAmount = lCharacterAmountCounter
	        
	        	If lCharacterAmountCounter > lMaxUnhandledCharacterAmount Then 
	        		sMaxUnhandledKeyword =  element
	        		iMaxUnhandledArticleAmount = iArticleAmountCounter
	        		lMaxUnhandledCharacterAmount = lCharacterAmountCounter
	        	End If
	        
	    		iKeywordComboNumber = iKeywordComboNumber + 1
	            iArticleAmountCounter = 0
    			lCharacterAmountCounter = 0
    	  
			End If
		
		Next element
		
		Redim Preserve HandledKeywordCombos(Ubound(HandledKeywordCombos)+1)
		HandledKeywordCombos(Ubound(HandledKeywordCombos)) = sMaxUnhandledKeyword
		
		Sheet3FirstColumnCell.string = sMaxUnhandledKeyword
		Sheet3SecondColumnCell.value = iMaxUnhandledArticleAmount		
	    Sheet3ThirdColumnCell.value =  lMaxUnhandledCharacterAmount
	    
	    iSheet3Row = iSheet3Row + 1
	    Sheet3FirstColumnCell = Sheet3.getCellByPosition(0,iSheet3Row)
    	Sheet3SecondColumnCell = Sheet3.getCellByPosition(1,iSheet3Row)
		Sheet3ThirdColumnCell = Sheet3.getCellByPosition(2,iSheet3Row)
		
		'may skip keywords handled in previous cycle
		iKeywordComboNumber = iKeyWordStartNumber
		   
	Loop
	
	
End Sub

'This procedure makes sheet with keywords, number of articles and number of characters

Sub MakeKeywordsSheet

	Dim Doc As Object
	Dim Sheet1 As Object
	Dim Sheet2 As Object
	Dim Sheet1ThirdColumnCell As Object
	Dim Sheet1FourthColumnCell As Object
	Dim Sheet2FirstColumnCell As Object
	Dim Sheet2SecondColumnCell As Object
	Dim Sheet2ThirdColumnCell As Object
    Dim Sheet2FourthColumnCell As Object
    
	Dim sKeywords() As String
	Dim sCellValue() As String ' support variable
	Dim element as Variant
	Dim element2 as Variant
	Dim iSheet1Row as Integer
    Dim iSheet2Row as Integer
    Dim iArticleAmountCounter as Integer
    Dim lCharacterAmountCounter as Long
    Dim iFirstMention as Integer
    
    
    Doc = ThisComponent
	Sheet1 = Doc.Sheets(0)
	Sheet2 = Doc.Sheets(1)
    sKeywords = GetKeywords()
    iSheet1Row = 1
    iSheet2Row = 0
    iArticleAmountCounter = 0
    lCharacterAmountCounter = 0
    iFirstMention = 0
    Sheet1ThirdColumnCell = Sheet1.getCellByPosition(2,iSheet1Row)
    Sheet1FourthColumnCell = Sheet1.getCellByPosition(3,iSheet1Row)
    Sheet2FirstColumnCell = Sheet2.getCellByPosition(0,iSheet2Row)
    Sheet2SecondColumnCell = Sheet2.getCellByPosition(1,iSheet2Row)
	Sheet2ThirdColumnCell = Sheet2.getCellByPosition(2,iSheet2Row)
    Sheet2FourthColumnCell = Sheet2.getCellByPosition(3,iSheet2Row)
    
    Sheet2FirstColumnCell.String = "Keyword"
    Sheet2SecondColumnCell.String = "Match number"
    Sheet2ThirdColumnCell.String = "Character amount"
    Sheet2FourthColumnCell.String = "First mention"
    
    iSheet2Row = 1
    Sheet2FirstColumnCell = Sheet2.getCellByPosition(0,iSheet2Row)
    Sheet2SecondColumnCell = Sheet2.getCellByPosition(1,iSheet2Row)
	Sheet2ThirdColumnCell = Sheet2.getCellByPosition(2,iSheet2Row)
    Sheet2FourthColumnCell = Sheet2.getCellByPosition(3,iSheet2Row)
        
    For each element in sKeywords
        If len(Cstr(element)) > 0 Then
			Sheet2FirstColumnCell.String = element
		
			Do Until strcomp(Sheet1FourthColumnCell.String, "End") = 0 OR IsEmpty(Sheet1FourthColumnCell)
	            sCellValue = Split(Sheet1FourthColumnCell.String,",") 		
	    		For Each element2 in sCellvalue
	    			If strcomp(element,Trim(element2)) = 0 Then
	    				'count number of articles
		    		   iArticleAmountCounter = iArticleAmountCounter+1
						'count number of characters
					   lCharacterAmountCounter = lCharacterAmountCounter + Sheet1ThirdColumnCell.Value
						'mark first mention 
						If iFirstMention = 0 Then iFirstMention = iSheet1Row
						'Keyword has been found and that is enough for inclusion
						Exit For 
					End If
				Next
				iSheet1Row = iSheet1Row + 1
	    		Sheet1ThirdColumnCell = Sheet1.getCellByPosition(2,iSheet1Row)
            	Sheet1FourthColumnCell = Sheet1.getCellByPosition(3,iSheet1Row)
			Loop
		End If
		
        Sheet2SecondColumnCell.value = iArticleAmountCounter		
	    Sheet2ThirdColumnCell.value =  lCharacterAmountCounter
	    Sheet2FourthColumnCell.value =  iFirstMention  
	    
	    iSheet2Row = iSheet2Row + 1
	    Sheet2FirstColumnCell = Sheet2.getCellByPosition(0,iSheet2Row)
    	Sheet2SecondColumnCell = Sheet2.getCellByPosition(1,iSheet2Row)
		Sheet2ThirdColumnCell = Sheet2.getCellByPosition(2,iSheet2Row)
		Sheet2FourthColumnCell = Sheet2.getCellByPosition(3,iSheet2Row)
		
		iSheet1Row = 1
    	Sheet1ThirdColumnCell = Sheet1.getCellByPosition(2,iSheet1Row)
    	Sheet1FourthColumnCell = Sheet1.getCellByPosition(3,iSheet1Row)   
	    iArticleAmountCounter = 0
    	lCharacterAmountCounter = 0
    	iFirstMention = 0
	      
	Next element   
    
End Sub


'First version works with empty combolist in the beginning. 

Private Function ArticleHasHandledKeywordCombo(ArticleKeywords As String,KeywordComboList() As String) 
	
	Dim element as Variant
	Dim returnvalue as boolean 
	
	returnvalue = false
	
	For Each element in KeywordComboList
		If SubWordList(element,ArticleKeywords) Then
			returnvalue = true
			Exit For
		End If
	Next

	ArticleHasHandledKeywordCombo = returnvalue
		
End Function 

'This propably useless function

'Private Function NotSubkeywords(KeywordCombo As String,KeywordComboList() As String) 
	
'	Dim element As Variant
'	Dim returnvalue As boolean 
	
'	returnvalue = false
	
'	For Each element in KeywordComboList
'		If SubWordList(KeywordCombo,element) Then
'			returnvalue = true
'			Exit For
'		End If
'	Next

'	NotSubkeywords = Not returnvalue
		
'End Function 

'Support function which investigates if Wordlist1 is a subwordlist of Wordlist2

Private Function SubWordList(Wordlist1 As String, Wordlist2 As String)
	
	Dim Wordlist1Splitted() As String
	Dim Wordlist2Splitted() As String
	Dim WordList1element As Variant
	Dim WordList2element As Variant
	Dim returnvalue As boolean
	Dim oneinclusion As boolean
	
	Wordlist1Splitted = Split(Wordlist1,",")
	Wordlist2Splitted = Split(Wordlist2,",")
    returnvalue = true
    oneinclusion = false
    For Each WordList1element in Wordlist1SPlitted
    	For Each WordList2element in Wordlist2Splitted
    		If strcomp(trim(WordList1element),trim(WordList2element)) = 0 Then 
    			oneinclusion = true
    			Exit For
    		End If
    	Next
    	If Not oneinclusion Then 
    		returnvalue = false
    		Exit For
    	End If
    	oneinclusion = false
    Next
 	SubWordList = returnvalue
End Function
   
'Support function which parses keywords.

Private Function GetKeywords() 'as String()
	Dim Doc As Object
	Dim Sheet As Object
	Dim FirstColumnCell As Object
	Dim ThirdColumnCell As Object
    Dim sFirstColumnString As String
    Dim sThirdColumnString as String
   
	Dim iRow As Integer            ' stores the current row number
	Dim iKeywordAmount As Integer 'stores amount of keywords
	Dim sKeywords() As String  ' array to store the cell values
	Dim sCellValue() As String ' support variable
	Dim element as variant
	
	Doc = ThisComponent
	Sheet = Doc.Sheets(0)
    iRow = 1
	FirstColumnCell = Sheet.getCellByPosition(0,iRow)
	ThirdColumnCell = Sheet.getCellByPosition(3,iRow)
	sFirstColumnString = FirstColumnCell.String
	sThirdColumnString = ThirdColumnCell.String
	iKeywordAmount = 0
	' Do Until loop to extract the value of each cell in column A
	' of the active Worksheet, as long as the first is not "End" or blank
	Do Until strcomp(sFirstColumnString, "End") = 0 OR IsEmpty(FirstColumnCell) 
	    ' Store the current cell in the CellValues array
		sCellValue = Split(sThirdColumnString,",")
	    For each element in sCellValue
	      If Not IsInArray(Trim(element),sKeywords) AND Len(Trim(element)) > 0 Then
	      	 iKeywordAmount = iKeywordAmount + 1
             ReDim Preserve sKeywords(1 To iKeywordAmount)
	      	 sKeywords(iKeywordAmount) = Trim(element)
	      End If
	    Next element   
	    iRow = iRow + 1
	    FirstColumnCell = Sheet.getCellByPosition(0,iRow)
		ThirdColumnCell = Sheet.getCellByPosition(3,iRow)
		sFirstColumnString = FirstColumnCell.String
	    sThirdColumnString = ThirdColumnCell.String
	Loop
	GetKeywords = sKeywords
End Function

Private Function IsInArray(valToBeFound As Variant, arr As Variant) As Boolean
'DEVELOPER: Ryan Wells (wellsr.com)
'DESCRIPTION: Function to check if a value is in an array of values
'INPUT: Pass the function a value to search for and an array of values of any data type.
'OUTPUT: True if is in array, false otherwise
Dim element As Variant
On Error GoTo IsInArrayError: 'array is empty
    For Each element In arr
        If element = valToBeFound Then
            IsInArray = True
            Exit Function
        End If
    Next element
Exit Function
IsInArrayError:
On Error GoTo 0
IsInArray = False
End Function

'Procedures for testing 

Sub TestArticleHasHandledKeywordCombo()
	Dim b As boolean
	Dim Wordlist(1 to 3) As String
	Wordlist(1) = "A,B,C,D,E"
	Wordlist(2) = "D,F"
	Wordlist(3) = "A,B,C,D"
	b = ArticleHasHandledKeywordCombo("A,B,C",WordList)
	MsgBox(b)
End Sub

Sub TestNotSubkeyWords
	Dim b as boolean 
	Dim Wordlist(1) As String
	Wordlist(0) = "Venäjä,Suomi"
	b = NotSubKeywords("Venäjä",WordList)
	MsgBox(b)
End Sub

Sub TestSubWordList
	Dim b as boolean
	b = SubWordList("A,B","B,C,A")
	MsgBox(b)
End Sub

Sub TestArticleHasHandledKeywordCombo2
	Dim b as boolean 
	Dim Wordlist(1) As String
	Wordlist(0) = "Venäjä,Suomi"
	b = ArticleHasHandledKeywordCombo("b",WordList)
	MsgBox(b)

End Sub

