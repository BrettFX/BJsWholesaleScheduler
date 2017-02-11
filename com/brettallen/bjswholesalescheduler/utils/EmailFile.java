package com.brettallen.bjswholesalescheduler.utils;

import com.brettallen.bjswholesalescheduler.gui.Menu;

import java.util.ArrayList;
import java.util.prefs.BackingStoreException;
import java.util.prefs.Preferences;

/**
 * Created by Brett Allen on 7/26/2016 at 2:43 PM.
 *
 *      The main function of this class is to permit the user to add and eraseEmail emails from a
 *      list. Instead of writing to a file at the root and requiring that file to be at the same
 *      location at the beginning of each session, this class eliminates having a separate file
 *      floating around outside of the program, enhancing organization.
 */
public class EmailFile
{
    private final Preferences emailPrefs;

    public EmailFile()
    {
        emailPrefs = Preferences.userRoot().node(this.getClass().getName());
    }

    public void add(String key, String value)
    {
        emailPrefs.put(key, value);
    }

    public ArrayList<String> getEmailList()
    {
        ArrayList<String> emailList = new ArrayList<String>();

        try
        {
            //Gets the key (name) of each email entry and it's corresponding email
            for(String key : emailPrefs.keys())
            {
                if(!emailPrefs.get(key, "").equals(" "))
                    emailList.add(key + Menu.DEFAULT_SEPARATOR + emailPrefs.get(key, ""));
            }
        }
        catch(BackingStoreException e)
        {
            e.printStackTrace();
        }

        return emailList;
    }

    public void eraseEmail(String key)
    {
        //Erase email by setting it to blank
        emailPrefs.put(key, " ");

        try
        {
            emailPrefs.flush();
        }
        catch(BackingStoreException e)
        {
            e.printStackTrace();
        }
    }
}
