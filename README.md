# Domain User & Group Automator Manual (DUGAM)
V2.1.1

This is a powershell tool that is used to create AD user and mass adjust group membership on AD, Entra and exchanged by copying from an existing user. The full details can be found in the [manual](ManualFiles\Manual.md)

## Tool Functions
The two primary functions you can select from are create a new user and Mass Group membership. It will then follow the prompts for each respective function.
### Create a New AD/Entra User Account:

- **Use Case:**

    Ideal for creating new AD user that will be synced up to Entra. It will create a password from the password.dict file and random amount of numbers, and special characters to reach the minimum length. If desired it will copy all AD and Entra groups via a schedule task from an existing user. 
- **Prompts:**

    Will prompt for the following: First Name*, SurName*, Title, Department, Employee ID, Manager, Office Location, and Phone Extension. Fields with * are required. The Manager Prompt will ask for name, it will then search AD to find an user that matches. The Office locations will prompt you to select from the list built in config file. If one of the office location info is blank (City, Office name, State, Postal Code) it will then prompt the user for it. If there is an typo or the OU in config file is not found it will error out and exit. Once the user is created it will prompt if you want to copy groups from an existing users. 
  
- **Group Copying:**
    
    It will prompt if you want to copy group memberships from an existing user. If selected to copy, it will copy all group memberships from another user in AD, Entra, and Exchange. The user running the tool must have permission to adjust membership on the groups.  Entra and Exchange groups are copied via a scheduled task to ensure the new account fully syncs.

- **Notification:**

    Once the setup is complete, the tool will notify the user’s manager and blind carbon copy the IT team with the account details and password. This can be turn off via the config file setting, sendManagerEmail. 

### 2. Mass Group Membership:

- **Use Case:**
    This tool is ideal for when a user switches departments, sites, or different positions and they no longer need to access items from their previous position. It will remove all existing groups and copy new group membership over from an existing user for AD, Exchange, and Entra. 

- **Prompts:**
    Will prompt for the Target User, Source User. It will search for those users in AD and have you select from the found results. It will prompt if you want to adjust AD or Entra, both can be selected on a single run of the tool. It will also prompt if you want to remove existing groups before coping over new groups. 

- **Remove Existing Memberships:**
    Clears all current group memberships for AD and/or Entra depending on user selection.

- **Copy New Memberships:**
    Copies group memberships from another user for AD and/or Entra. User running the tool must have permission to adjust group membership.
    
