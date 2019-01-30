// initialiZe variables... 
var ppcIE=((navigator.appName == "Microsoft Internet Explorer") || ((navigator.appName == "Netscape") && (parseInt(navigator.appVersion)==5)));
var ppcNN6=((navigator.appName == "Netscape") && (parseInt(navigator.appVersion)==5));
//var ppcIE=(navigator.appName == "Microsoft Internet Explorer");
var ppcNN=((navigator.appName == "Netscape")&&(document.layers));
var ppcX = 4;
var ppcY = 4;
var cal_height = 130;
var IsCalendarVisible;
var calfrmName;
var maxYearList;
var minYearList;
var todayDate = new Date; 
var curDate = new Date; 
var curImg;
var curDateBox;
var minDate = new Date;
var maxDate = new Date;
var hideDropDowns;
var IsUsingMinMax;
var FuncsToRun;
var img_del;
var img_close;

img_del=new Image();
img_del.src="/includes/calendar_files/btn_del_small.gif";
img_close=new Image();
img_close.src="/includes/calendar_files/btn_close_small.gif";

minYearList=todayDate.getFullYear()-10;
maxYearList=todayDate.getFullYear()+10;
IsCalendarVisible=false;

img_Date_UP=new Image();
img_Date_UP.src="/includes/calendar_files/icon_calendar.gif";

img_Date_OVER=new Image();
img_Date_OVER.src="/includes/calendar_files/icon_calendar.gif";

img_Date_DOWN=new Image();
img_Date_DOWN.src="calendar_files/icon_calendar.gif";
function calSwapImg(whatID, NewImg,override) {
    if (document.images) {
     if (!( IsCalendarVisible && override )) {
        document.images[whatID].src = eval(NewImg + ".src");
     }
    }
    window.status=' ';
    return true;
}

function getOffsetLeft (el) {
    var ol = el.offsetLeft;
    while ((el = el.offsetParent) != null)
        ol += el.offsetLeft;
    return ol;
}

function getOffsetTop (el) {
    var ot = el.offsetTop;
    while((el = el.offsetParent) != null)
        ot += el.offsetTop;
    return ot;
}

function showCalendar(frmName, dteBox,btnImg, hideDrops, MnDt, MnMo, MnYr, MxDt, MxMo, MxYr,runFuncs, position, xpos) {
	if(position == null) {
		// 0 = below, 1 = above
		position = 0;
	}
    hideDropDowns = hideDrops;
    FuncsToRun = runFuncs;
    calfrmName = frmName;
    if (IsCalendarVisible) {
        hideCalendar();
    }
    else {
        if (document.images['calbtn1']!=null ) document.images['calbtn1'].src=img_del.src;
        if (document.images['calbtn2']!=null ) document.images['calbtn2'].src=img_close.src;
        
        if (hideDropDowns) {toggleDropDowns('hidden');}
        if ((MnDt!=null) && (MnMo!=null) && (MnYr!=null) && (MxDt!=null) && (MxMo!=null) && (MxYr!=null)) {
            IsUsingMinMax = true;
            minDate.setDate(MnDt);
            minDate.setMonth(MnMo-1);
            minDate.setFullYear(MnYr);
            maxDate.setDate(MxDt);
            maxDate.setMonth(MxMo-1);
            maxDate.setFullYear(MxYr);
        }
        else {
            IsUsingMinMax = false;
        }
        
        curImg = btnImg;
        curDateBox = dteBox;
		
        if ( ppcIE ) {
            if(xpos != null) {
				if(xpos < 0) {
					screenW = document.body.clientWidth;
					// screenH = document.body.clientHeight;
					ppcX = screenW + xpos;
				} else {
					ppcX = xpos;
				}
			} else {
				ppcX = getOffsetLeft(document.images[btnImg]);
			}

			if(position == 1)     
            	ppcY = getOffsetTop(document.images[btnImg]) - document.images[btnImg].height - cal_height;
			else
				ppcY = getOffsetTop(document.images[btnImg]) + document.images[btnImg].height;
			
        } else if (ppcNN || ppcNN6){
            if(xpos != null) {
				if(xpos < 0) {
					// screenW = document.body.clientWidth;
					// screenH = document.body.clientHeight;
					screenW = window.innerWidth;
					// screenH = window.innerHeight;
					ppcX = screenW + xpos;
				} else {
					ppcX = xpos;
				}
			} else {
				ppcX = document.images[btnImg].x;
			}
	 
			if(position == 1)  
				ppcY = document.images[btnImg].y - document.images[btnImg].height - cal_height;
			else
				ppcY = document.images[btnImg].y + document.images[btnImg].height;
        }

		// domlay('popupcalendar',1,ppcX,ppcY,Calendar(todayDate.getMonth(),todayDate.getFullYear())); 
		
		// if(monthfield == null) 
		//	monthvalue = todayDate.getMonth();
			
		// if(yearfield == null) 
		//	yearvalue = todayDate.getMonth();	
			
        // domlay('popupcalendar',1,ppcX,ppcY,Calendar(curDate.getMonth(),curDate.getFullYear(),curDate.getDate()));
		domlay('popupcalendar',1,ppcX,ppcY,Calendar(curDate.getMonth(),curDate.getFullYear()));

        IsCalendarVisible = true;
    }
}

function toggleDropDowns(showHow){
    var i; var j;
    for (i=0;i<document.forms.length;i++) {
        for (j=0;j<document.forms[i].elements.length;j++) {
            if (document.forms[i].elements[j].tagName == "SELECT") {
                if (document.forms[i].name != "Cal")
                    document.forms[i].elements[j].style.visibility=showHow;
            }
        }
    }
}

function hideCalendar(){
    domlay('popupcalendar',0,ppcX,ppcY);
    calSwapImg(curImg, 'img_Date_UP');    
    IsCalendarVisible = false;
    if (hideDropDowns) {toggleDropDowns('visible');}
}

function calClick() {
        window.focus();
}

/* function get_height(id) {
	if (document.layers) {
		return document.layers[''+id+''].height;
	} else if (document.all) {
		alert("IE");
		alert('height: '+document.all[''+id+''].style.height+'');
		return document.all[''+id+''].style.height;
	} else if (document.getElementById) {
		return document.getElementById(''+id+'').style.height;               
    } else {
		return 0;
	}
} */

function domlay(id,trigger,lax,lay,content,zindex) {
    /*
     * Cross browser Layer visibility / Placement Routine
     * Done by Chris Heilmann (mail@ichwill.net)
     * Feel free to use with these lines included!
     * Created with help from Scott Andrews.
     * The marked part of the content change routine is taken
     * from a script by Reyn posted in the DHTML
     * Forum at Website Attraction and changed to work with
     * any layername. Cheers to that!
     * Welcome DOM-1, about time you got included... :)
     */
    // Layer visible
	if(zindex == null) {
		zindex = 99;
	}
    if (trigger=="1"){
        // if (document.layers) document.layers[''+id+''].visibility = "show"
        // else if (document.all) document.all[''+id+''].style.visibility = "visible"
        // else if (document.getElementById) document.getElementById(''+id+'').style.visibility = "visible"                
        // }
        if (document.layers) {
			document.layers[''+id+''].visibility = "show";
			document.layers[''+id+''].zIndex = zindex;
		} else if (document.all) {
			document.all[''+id+''].style.visibility = "visible";
			document.all[''+id+''].style.zIndex = zindex;
		} else if (document.getElementById) {
			document.getElementById(''+id+'').style.visibility = "visible";  
			document.getElementById(''+id+'').style.zIndex = zindex;                
        }
	}
    // Layer hidden
    else if (trigger=="0"){
        if (document.layers) document.layers[''+id+''].visibility = "hide"
        else if (document.all) document.all[''+id+''].style.visibility = "hidden"
        else if (document.getElementById) document.getElementById(''+id+'').style.visibility = "hidden"             
        }
    // Set horizontal position  
    if (lax){
        if (document.layers){document.layers[''+id+''].left = lax}
        else if (document.all){document.all[''+id+''].style.left=lax}
        else if (document.getElementById){document.getElementById(''+id+'').style.left=lax+"px"}
        }
    // Set vertical position
    if (lay){
        if (document.layers){document.layers[''+id+''].top = lay}
        else if (document.all){document.all[''+id+''].style.top=lay}
        else if (document.getElementById){document.getElementById(''+id+'').style.top=lay+"px"}
        }
    // change content

    if (content){
    if (document.layers){
        sprite=document.layers[''+id+''].document;
        // add father layers if needed! document.layers[''+father+'']...
        sprite.open();
        sprite.write(content);
        sprite.close();
        }
    else if (document.all) document.all[''+id+''].innerHTML = content;  
    else if (document.getElementById){
        //Thanx Reyn!
        rng = document.createRange();
        el = document.getElementById(''+id+'');
        rng.setStartBefore(el);
        htmlFrag = rng.createContextualFragment(content)
        while(el.hasChildNodes()) el.removeChild(el.lastChild);
        el.appendChild(htmlFrag);
        // end of Reyn ;)
        }
    }
}

function Calendar(whatMonth,whatYear,whatDay) {
    var output = '';
    var datecolwidth;
    var startMonth;
    var startYear;
    startMonth=whatMonth;
    startYear=whatYear;

    curDate.setMonth(whatMonth);
    curDate.setFullYear(whatYear);
    if(whatDay != null) {
		curDate.setDate(whatDay);
	}
	// } else {
	//	curDate.setDate(todayDate.getDate());
	// }

    if (ppcNN6) {
        output += '<form name="Cal"><table width="185" border="2" class="cal-Table" cellspacing="0" cellpadding="0"><tr>';
    }
    else {
        output += '<table width="185" border="2" class="cal-Table" cellspacing="0" cellpadding="0"><form name="Cal"><tr>';
    }
     
    output += '<td class="cal-HeadCell" align="center" width="100%"><a href="javascript:clearDay();"><img name="calbtn1" src="/intranet/includes/calendar_files/btn_del_small.gif" border="0" width="12" height="10"></a>  <a href="javascript:scrollMonth(-1);" class="cal-DayLink"><</a> <SELECT class="cal-TextBox" NAME="cboMonth" onChange="changeMonth();">';
    for (month=0; month<12; month++) {
        if (month == whatMonth) output += '<OPTION VALUE="' + month + '" SELECTED>' + names[month] + '<\/OPTION>';
        else                output += '<OPTION VALUE="' + month + '">'          + names[month] + '<\/OPTION>';
    }

    output += '<\/SELECT><SELECT class="cal-TextBox" NAME="cboYear" onChange="changeYear();">';

    for (year=minYearList; year<maxYearList; year++) {
        if (year == whatYear) output += '<OPTION VALUE="' + year + '" SELECTED>' + year + '<\/OPTION>';
        else              output += '<OPTION VALUE="' + year + '">'          + year + '<\/OPTION>';
    }

    output += '<\/SELECT>&nbsp;<a href="javascript:scrollMonth(1);" class="cal-DayLink">&gt;</a>&nbsp;&nbsp;<a href="javascript:hideCalendar();"><img name="calbtn2" src="/includes/calendar_files/btn_close_small.gif" border="0" width="12" height="10"></a><\/td><\/tr><tr><td width="100%" align="center">';

    firstDay = new Date(whatYear,whatMonth,1);
    startDay = firstDay.getDay();
	
    if (((whatYear % 4 == 0) && (whatYear % 100 != 0)) || (whatYear % 400 == 0))
         days[1] = 29;
    else
         days[1] = 28;

    output += '<table width="185" cellspacing="1" cellpadding="2" border="0"><tr>';

    for (i=0; i<7; i++) {
        if (i==0 || i==6) {
            datecolwidth="15%"
        }
        else
        {
            datecolwidth="14%"
        }
        output += '<td class="cal-HeadCell" width="' + datecolwidth + '" align="center" valign="middle">'+ dow[i] +'<\/td>';
    }
    
    output += '<\/tr><tr>';

    var column = 0;
    var lastMonth = whatMonth - 1;
    var lastYear = whatYear;
    if (lastMonth == -1) { lastMonth = 11; lastYear=lastYear-1;}

    for (i=0; i<startDay; i++, column++) {
        output += getDayLink((days[lastMonth]-startDay+i+1),true,lastMonth,lastYear);
    }

    for (i=1; i<=days[whatMonth]; i++, column++) {
        output += getDayLink(i,false,whatMonth,whatYear);
        if (column == 6) {
            output += '<\/tr><tr>';
            column = -1;
        }
    }
    
    var nextMonth = whatMonth+1;
    var nextYear = whatYear;
    if (nextMonth==12) { nextMonth=0; nextYear=nextYear+1;}
    
    if (column > 0) {
        for (i=1; column<7; i++, column++) {
            output +=  getDayLink(i,true,nextMonth,nextYear);
        }
        output += '<\/tr><\/table><\/td><\/tr>';
    }
    else {
        output = output.substr(0,output.length-4); // remove the <tr> from the end if there's no last row
        output += '<\/table><\/td><\/tr>';
    }
    
    if (ppcNN6) {
        output += '<\/table><\/form>';
    }
    else {
        output += '<\/form><\/table>';
    }
    curDate.setDate(1);
    curDate.setMonth(startMonth);
    curDate.setFullYear(startYear);
    return output;
}

function getDayLink(linkDay,isGreyDate,linkMonth,linkYear) {
    var templink;
    if (!(IsUsingMinMax)) {
        if (isGreyDate) {
            templink='<td align="center" class="cal-GreyDate">' + linkDay + '<\/td>';
        }
        else {
            if (isDayToday(linkDay)) {
                templink='<td align="center" class="cal-DayCell">' + '<a class="cal-TodayLink" onmouseover="self.status=\' \';return true" href="javascript:changeDay(' + linkDay + ');">' + linkDay + '<\/a>' +'<\/td>';
            }
            else {
                templink='<td align="center" class="cal-DayCell">' + '<a class="cal-DayLink" onmouseover="self.status=\' \';return true" href="javascript:changeDay(' + linkDay + ');">' + linkDay + '<\/a>' +'<\/td>';
            }
        }
    }
    else {
        if (isDayValid(linkDay,linkMonth,linkYear)) {

            if (isGreyDate){
                templink='<td align="center" class="cal-GreyDate">' + linkDay + '<\/td>';
            }
            else {
                if (isDayToday(linkDay)) {
                    templink='<td align="center" class="cal-DayCell">' + '<a class="cal-TodayLink" onmouseover="self.status=\' \';return true" href="javascript:changeDay(' + linkDay + ');">' + linkDay + '<\/a>' +'<\/td>';
                }
                else {
                    templink='<td align="center" class="cal-DayCell">' + '<a class="cal-DayLink" onmouseover="self.status=\' \';return true" href="javascript:changeDay(' + linkDay + ');">' + linkDay + '<\/a>' +'<\/td>';
                }
            }
        }
        else {
            templink='<td align="center" class="cal-GreyInvalidDate">'+ linkDay + '<\/td>';
        }
    }
    return templink;
}

function isDayToday(isDay) {
    if ((curDate.getFullYear() == todayDate.getFullYear()) && (curDate.getMonth() == todayDate.getMonth()) && (isDay == todayDate.getDate())) {
        return true;
    }
    else {
        return false;
    }
}

function isDayValid(validDay, validMonth, validYear){
    curDate.setDate(validDay);
    curDate.setMonth(validMonth);
    curDate.setFullYear(validYear);
    
    if ((curDate>=minDate) && (curDate<=maxDate)) {
        return true;
    }
    else {
        return false;
    }
}

function padout(number) { return (number < 10) ? '0' + number : number; }

function clearDay() {
    // eval('document.' + calfrmName + '.' + curDateBox + '.value = \'\'');
    hideCalendar();
    if (FuncsToRun!=null)
        eval(FuncsToRun); 
}

function changeDay(whatDay) {
    curDate.setDate(whatDay);
    // eval('document.' + calfrmName + '.' + curDateBox + '.value = "'+ padout(curDate.getDate()) + '-' + names[curDate.getMonth()] + '-' + curDate.getFullYear() + '"');
    
	hideCalendar();
    if (FuncsToRun!=null)
        eval(FuncsToRun);
}

function scrollMonth(amount) {
    var monthCheck;
    var yearCheck;
    
    if (ppcIE) {
        monthCheck = document.forms["Cal"].cboMonth.selectedIndex + amount;
    }
    else if (ppcNN) {
        monthCheck = document.popupcalendar.document.forms["Cal"].cboMonth.selectedIndex + amount;    
    }
    if (monthCheck < 0) {
        yearCheck = curDate.getFullYear() - 1;
        if ( yearCheck < minYearList ) {
            yearCheck = minYearList;
            monthCheck = 0;
        }
        else {
            monthCheck = 11;
        }
        curDate.setFullYear(yearCheck);
    }
    else if (monthCheck >11) {
        yearCheck = curDate.getFullYear() + 1;
        if ( yearCheck > maxYearList-1 ) {
            yearCheck = maxYearList-1;
            monthCheck = 11;
        }
        else {
            monthCheck = 0;
        }      
        curDate.setFullYear(yearCheck);
    }
    
    if (ppcIE) {
        curDate.setMonth(document.forms["Cal"].cboMonth.options[monthCheck].value);
    }
    else if (ppcNN) {
        curDate.setMonth(document.popupcalendar.document.forms["Cal"].cboMonth.options[monthCheck].value );
    }
    domlay('popupcalendar',1,ppcX,ppcY,Calendar(curDate.getMonth(),curDate.getFullYear()));
}

function changeMonth() {

    if (ppcIE) {        
        curDate.setMonth(document.forms["Cal"].cboMonth.options[document.forms["Cal"].cboMonth.selectedIndex].value);
        domlay('popupcalendar',1,ppcX,ppcY,Calendar(curDate.getMonth(),curDate.getFullYear()));
    }
    else if (ppcNN) {

        curDate.setMonth(document.popupcalendar.document.forms["Cal"].cboMonth.options[document.popupcalendar.document.forms["Cal"].cboMonth.selectedIndex].value);
        domlay('popupcalendar',1,ppcX,ppcY,Calendar(curDate.getMonth(),curDate.getFullYear()));
    }

}

function changeYear() {
    if (ppcIE) {

        curDate.setFullYear(document.forms["Cal"].cboYear.options[document.forms["Cal"].cboYear.selectedIndex].value);
        domlay('popupcalendar',1,ppcX,ppcY,Calendar(curDate.getMonth(),curDate.getFullYear()));
    }
    else if (ppcNN) {

        curDate.setFullYear(document.popupcalendar.document.forms["Cal"].cboYear.options[document.popupcalendar.document.forms["Cal"].cboYear.selectedIndex].value);
        domlay('popupcalendar',1,ppcX,ppcY,Calendar(curDate.getMonth(),curDate.getFullYear()));
    }
}

function makeArray0() {
    for (i = 0; i<makeArray0.arguments.length; i++)
        this[i] = makeArray0.arguments[i];
}

function resetDays(selObjM, selObjD, selObjY) {
	if(!selObjM || !selObjD) {
		return;
	}
	
	month = selObjM.selectedIndex + 1;
	dayInd = selObjD.selectedIndex;
	if(selObjY == null) {
		year = get_valid_year(month);	
	} else {
		year = selObjY.options[selObjY.selectedIndex].value;
	}
	// clear day select
	daylength = selObjD.options.length
	for (i = daylength; i > 0; i--) {
		selObjD.options[i] = null;
	}
	// recreate day options
	var lastDay = getDaysInMonth(month, year);
	// confirm(lastDay);
	for (i = 1; i <= lastDay; i++) {
		selObjD.options[i-1] = new Option(i,i);
	}
	// set selected index
	selObjD.selectedIndex = dayInd;

	// update_calendarfields(selObjM, selObjD, selObjY);
}

function getDaysInMonth(month, year)  {
    var days;
    // var month = calDate.getMonth()+1;
    // var year  = calDate.getFullYear();

    // RETURN 31 DAYS
    if (month==1 || month==3 || month==5 || month==7 || month==8 ||
        month==10 || month==12)  {
        days=31;
    }
    // RETURN 30 DAYS
    else if (month==4 || month==6 || month==9 || month==11) {
        days=30;
    }
    // RETURN 29 DAYS
    else if (month==2)  {
        if (isLeapYear(year)) {
            days=29;
        }
        // RETURN 28 DAYS
        else {
            days=28;
        }
    }
    return (days);
}


// CHECK TO SEE IF YEAR IS A LEAP YEAR
function isLeapYear (Year) {
    if (((Year % 4)==0) && ((Year % 100)!=0) || ((Year % 400)==0)) {
        return (true);
    } else {
        return (false);
    }
}

function get_valid_year(valid_month, valid_day) {
	if(valid_day == null) {
		valid_day = 1;
	}
	var current_Date = new Date(); 
	current_month = current_Date.getMonth();
	current_day = current_Date.getDate();
	if(current_month > valid_month || (current_month == valid_month && current_day > valid_day)) {
		valid_year = current_Date.getFullYear() + 1;
	} else {
		valid_year = current_Date.getFullYear();
	}
	return valid_year;
}

function update_datefields(monthfd, dayfd, yearfd) {
    if(yearfd != null) {
		yearfd.selectedIndex = curDate.getFullYear() - minDate.getFullYear();
	}
	monthfd.selectedIndex = curDate.getMonth();
	resetDays(monthfd, dayfd, yearfd);
	dayfd.selectedIndex = curDate.getDate() - 1;
	return true;
}

function update_calendarfields(monthfd, dayfd, yearfd) {
	curDate.setMonth(monthfd.selectedIndex);
	if(yearfd != null) {
		curDate.setFullYear(yearfd.options[yearfd.selectedIndex].value);
	} else {
		curDate.setFullYear(todayDate.getFullYear());
	}
	day = dayfd.selectedIndex + 1
	curDate.setDate(day);
	return true;
}


var names     = new makeArray0('Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec');
var days      = new makeArray0(31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31);
var dow       = new makeArray0('S','M','T','W','T','F','S');
