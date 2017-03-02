package application;

import java.util.Date;

public class WeekRange {
	Date startWeek;
	Date endWeek;
	public int getMonth() {
		return month;
	}

	public void setMonth(int month) {
		this.month = month;
	}

	public int getMonthWeek() {
		return monthWeek;
	}

	public void setMonthWeek(int monthWeek) {
		this.monthWeek = monthWeek;
	}

	int month;
	int monthWeek;
	
	public WeekRange(){	
	}
	
	public Date getStartWeek(){
		return(startWeek);
	}
	
	public void setStartWeek(Date sw){
		this.startWeek = sw;
	}
	
	public Date getEndWeek(){
		return(endWeek);
	}
	
	public void setEndWeek(Date ew){
		this.endWeek = ew;
	}
	
}
