package com.brettallen.bjswholesalescheduler.utils;

/*
 * Created By: Brett Allen on Oct 22, 2015 at 1:44:50 PM
 */

import com.brettallen.bjswholesalescheduler.BJsWholesaleScheduler;

public class Shift
{
	private String separator;

    private final boolean specialShift;

	public enum EmployeeType {NON_MANAGER, OPENING_MANAGER, MID_SHIFT_MANAGER, CLOSING_MANAGER}
	private EmployeeType employeeType = EmployeeType.NON_MANAGER;

	public final String employee;
	public String position;
	public final String startTime;
	public final String endTime;
	public String day;

	public Shift(String employee, String position, String startTime, String endTime, int day, boolean isSpecial)
	{
		separator = "";

        specialShift = isSpecial;

		this.employee = employee.trim();
		this.position = position.trim();
		this.startTime = startTime.trim();
		this.endTime = endTime.trim();
		
		switch(day)
		{
		case 0:
			this.day = "Sunday";
			break;
		case 1:
			this.day = "Monday";
			break;
		case 2:
			this.day = "Tuesday";
			break;
		case 3:
			this.day = "Wednesday";
			break;
		case 4:
			this.day = "Thursday";
			break;
		case 5:
			this.day = "Friday";
			break;
		case 6:
			this.day = "Saturday";
		default:
			break;
		}

		//Handle the managerial shifts accordingly
		if(position.contains("Manager") || position.contains("Mgr"))
			determineManagerType();
	}

	/*Determines the manager type based the manager's shift time*/
	private void determineManagerType()
	{
		//Determine if the shift starts with a number (specific shift time)
		if(BJsWholesaleScheduler.startsWithNumber(startTime))
		{
			//Separate the start and end times for managers with specific times
			String newEndTime = startTime.split("-")[1];

			//Determine which block the manager falls in
			//Closing (endTime is greater than 9:00p)
			if(BJsWholesaleScheduler.getMilitaryTime(newEndTime) > 2100)
				employeeType = EmployeeType.CLOSING_MANAGER;
			//Mid-shift (endTime is greater than 5:00p)
			else if(BJsWholesaleScheduler.getMilitaryTime(newEndTime) > 1700)
				employeeType = EmployeeType.MID_SHIFT_MANAGER;
			//Opening (endTime is greater than 12:00p)
			else if(BJsWholesaleScheduler.getMilitaryTime(newEndTime) > 1200)
				employeeType = EmployeeType.OPENING_MANAGER;
		}
		else
		{
			//Set the type of employee based on if the employee is a manager or not
			if(startTime.contains("Open") || startTime.contains("Early"))
				employeeType = EmployeeType.OPENING_MANAGER;
			else if(startTime.contains("Mid"))
				employeeType = EmployeeType.MID_SHIFT_MANAGER;
			else if(startTime.contains("Close"))
				employeeType = EmployeeType.CLOSING_MANAGER;
		}
	}

	public void setSeparator(@SuppressWarnings("SameParameterValue") String s)
	{
		separator = s;
	}

	public boolean isSpecial()
    {
        return specialShift;
    }

    public EmployeeType getEmployeeType()
	{
		return employeeType;
	}

	public String getSeparator()
	{
		return separator;
	}

	public String getShiftTime()
	{
	    //Simply return the start time if the employee has a special shift
	    if(specialShift)
	        return startTime;

		String[] truncStart = new String[2];
		
		if(startTime.contains("a"))
			truncStart = startTime.split("a");
		else if(startTime.contains("p"))
			truncStart = startTime.split("p");
		
		return truncStart[0] + "-" + endTime;
	}
	
	public String getName()
	{
		String nickName;
		
		if(employee.split(",")[1].trim().contains("Christopher"))
		{
			nickName = "Chris";
			
			return nickName + " " + employee.split(",")[0].trim().charAt(0);
		}		
		else if(employee.split(",")[1].trim().contains("Mekdes"))
		{
			nickName = "Mimi";
			
			return nickName + " " + employee.split(",")[0].trim().charAt(0);
		}
		
		return employee.split(",")[1].trim() + " " + employee.split(",")[0].trim().charAt(0);
	}
	
	public String getFullName()
	{
		return employee.split(",")[1].trim() + " " + employee.split(",")[0].trim();
	}
	
	/**Prepends an asterisk to the name to indicate special*/
	public String getSpecialCaseName() {
		String name = employee.split(",")[1].trim() + " " + employee.split(",")[0].trim().charAt(0);
		return "* " + name;
	}

	@Override
	public String toString() 
	{
	    if(specialShift)
        {
            return "Day: " + day + "\nEmployee: " + employee + "\nPosition: " + position +
                    "\nTime: " + startTime + "\n";
        }

		return "Day: " + day + "\nEmployee: " + employee + "\nPosition: " + position +
				"\nTime: " + startTime + " - " + endTime + "\n";
	}
}