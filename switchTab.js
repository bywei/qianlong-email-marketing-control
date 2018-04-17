	var debug = false ;
	function switchTab (tab)
	{	event.returnValue = false ;	 
		var tabName = getTabGroupName (tab.id) ;
		if (tabName == '')
		{	alert ("No tabName for tab [" + tab.id + "]") ;
			return ;
		}
		var index = 1 ;
		while (true)
		{	var tabTitle = eval ("document.all ('" + tabName + '_' + index + "')") ;
			if (tabTitle == undefined)
				break ;
			deactiveTabTitle (tabTitle , tabName) ;
			var tabContent = eval ("document.all ('" + tabName + '_' + index + '_content' + "')") ;
			if (tabContent != undefined)
				tabContent.style.display = "NONE" ;
			index ++ ;
		}
		if (debug)
			alert ("Find " + (index - 1) + " tab title(s) for TabName [" + tabName + "]") ;
		activeTabTitle (tab , tabName) ;
		var tabContent = eval ("document.all ('" + tab.id + '_content' + "')") ;
		if (tabContent != undefined)
			tabContent.style.display = "BLOCK" ;
	}	
	function getTabGroupName (tabId)
	{
		if (tabId == '' || tabId == undefined)
		{
			alert ("tabId is NULL! [" + tabId + "]") ;
			return ;
		}
		var i = tabId.lastIndexOf ('_') ;
		if (i <= 1)
			return '' ;
		return tabId.substr (0 , i) ;
	}
	function deactiveTabTitle (tab , tabName)
	{	//	tab.className = tabName + "_off" ;
		classname=tab.className;
		tab.className = classname.replace("on","off") ;
	}
	
	function activeTabTitle (tab , tabName)
	{	//	tab.className = tabName + "_on" ;
		classname=tab.className;
		tab.className = classname.replace("off","on") ;
	}
// -->	