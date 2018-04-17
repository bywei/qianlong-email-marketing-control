function HMRolling()
{
	this.version = "1.0.6.0227";
	this.name = "HMRolling";	
	this.HeaderText = "";		
	this.FooterText = "";		

	this.item = new Array();	
	this.itemcount = 0;
	this.mitem = new Array();	
	this.mitemcount = 0;
	this.curmitemidx = 0;
	this.itemdivision = 5;		
	
	this.AutoRepli = true;		
	this.AutoRolling = true;	

	this.currentspeed = 0;
	this.scrollspeed = 1000;	
	this.pausedelay = 1000;

	this.pausemouseover = false;
	this.stop = false;

	this.add = function()
	{
		var text = arguments[0];
		this.item[this.itemcount] = text;
		this.itemcount = this.itemcount + 1;
	}//add
	this.start = function()
	{	

		if(this.itemcount == "")
		{
			return;
		}

		var lp;
		if(this.itemcount % this.itemdivision == 0)
		{
			lp = parseInt(this.itemcount / this.itemdivision);
		}
		else
		{
			lp = parseInt(this.itemcount / this.itemdivision) + 1;
		}

		var curitemidx = 0;
		var contents = "";
		for(i = 0 ; i < lp ; i++)
		{
			for(j = 0 ; j < this.itemdivision ; j ++)
			{

				if(curitemidx >= this.itemcount)
				{
					if (this.AutoRepli == false)
					{
						break;
					}
					else
					{
						curitemidx = 0;
					}
				}
				contents += this.item[curitemidx];
				curitemidx++;				
			}
			this.mitem[i] = contents;
			contents = ""
			this.mitemcount = this.mitemcount + 1;
		}
		this.display();
		this.currentspeed = this.scrollspeed;
		this.stop = true;
		setTimeout(this.name+'.rolling()',this.currentspeed);
		window.setTimeout(this.name+".stop = false", this.pausedelay);
	}//start
	this.display = function()
	{
		document.write('<div id="'+this.name+'" style="position:relative; overflow:hidden; " OnMouseOver="'+this.name+'.onmouseover(); " OnMouseOut="'+this.name+'.onmouseout(); ">');
		document.write(this.HeaderText);
		document.write(this.mitem[0]);
		document.write(this.FooterText);
		document.write('</div>');
	}//display
	this.rolling = function()
	{
		if(this.AutoRolling == false) return;
		
		if (this.stop == false)this.next();
		window.setTimeout(this.name+".rolling()",this.scrollspeed);
	}//rolling
	this.next = function() 
	{
		if(this.curmitemidx >= this.mitemcount - 1)
		{
			this.curmitemidx = 0;
		}
		else
		{
			this.curmitemidx++;
		}
		var obj = document.getElementById(this.name);
		var contents = "";
		contents = this.HeaderText;
		contents += this.mitem[this.curmitemidx];
		contents += this.FooterText;
		obj.innerHTML = contents;

	}//next
	this.prev = function() 
	{
		if((this.curmitemidx) <= 0)
		{
			this.curmitemidx = this.mitemcount - 1;
		}
		else
		{
			this.curmitemidx--;
		}
		var obj = document.getElementById(this.name);
		var contents = "";
		contents = this.HeaderText;
		contents += this.mitem[this.curmitemidx];
		contents += this.FooterText;
		obj.innerHTML = contents;
	}//prev
	this.unext = function () 
	{
		this.onmouseover();
		this.next();
		window.setTimeout(this.name+".onmouseout()",this.pausedelay);
	}//unext
	this.uprev = function ()
	{
		this.onmouseover();
		this.prev();
		window.setTimeout(this.name+".onmouseout()",this.pausedelay);
	}//uprev
	this.onmouseover = function()
	{
		if(this.pausemouseover) this.stop = true;
	}//onmouseover
	this.onmouseout = function()
	{
		if(this.pausemouseover) this.stop = false;
	}//onmouseout
}//HMScroll
-->
