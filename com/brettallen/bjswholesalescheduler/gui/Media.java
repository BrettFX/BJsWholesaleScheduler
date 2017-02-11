/*
*Created By: Brett Allen on May 20, 2016 at 11:55:03 AM
*	
*	Run/play media using the host computer's default media application
*
*/

package com.brettallen.bjswholesalescheduler.gui;

import java.awt.Desktop;
import java.io.File;

class Media
{
	private final File mediaFile;

	boolean fileExists;
	
	Media(String path)
	{
		mediaFile = new File(path);
	}
	
	void playVideo()
	{
		try
		{			
			if(mediaFile.exists())
			{
				fileExists = true;

				if(Desktop.isDesktopSupported())
				{
					Desktop.getDesktop().open(mediaFile);
				}
				else
				{
					System.err.println("Cannot open file; File not supported by desktop");
				}
			}
			else
				fileExists = false;
		}
		catch(Exception e)
		{
			System.err.println("Media File Not Found!");
			e.printStackTrace();
		}
	}
}
