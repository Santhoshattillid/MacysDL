
Macys DL

This library helps you to pull the content from sharepoint site using web services. This will provides you object model so that you can integrate with your application easily. 

Please go through with the following for integration

1. Download the source and build for MacysSPDL.dll, otherwise you can simply download from bin\debug folder.

2. Add reference to your project and use the following sample code 


// define the column definitions instance with required parameters
 
var columnDefinition = new ColumnDefinition(	
    "content", // content field name from list
    "Topic", // Topic field name from list
    "Location", // Location field name from list
    "Job Code", // JobCode field name from list
    "Sub Topic" // Sub Topic field name from list
);

columnDefinition.JobCodeColumnType = "Text"; // define JobCode column type, default is 'Text'
columnDefinition.LocationColumnType = "Text"; // define Location column type, default is 'Text'
columnDefinition.SubTopicColumnType = "Text"; // define Sub Topic column type, default is 'Text'
columnDefinition.TopicColumnType = "Text"; // define Topic column type, default is 'Text'

// instantiate the macyDL object
    
var macysDL = new MacysDL(
    @"http://tspsrvr:666", // sharepoint site URL
    "LabourEssentials", // List name
    System.Net.CredentialCache.DefaultNetworkCredentials, // Credentials used to connect the sharepoint web services
    columnDefinition // Column definition instance
    );

// Access your required method for content

// method to get content by filtering with topic, jobcode and location.
var content = macysDL.GetContents(
    "Topic1", // Topic phrase to filter the records
    "71001", //  Jobcode to filter the records
    "Chennai" // Location to filter the records
    );

// method to get content by filtering with jobcode and location. It will returns multiple contents using Generic Lists object
List<string> contents = macysDL.GetAllContents(
    "71001", //  Jobcode to filter the records
    "Chennai" // Location to filter the records
    );

// method to get topic and sub topic.  It will returns multiple contents using Generic Lists object
List<string> topicAndSubtopics = macysDL.GetTopicsAndSubTopics(
    "71001", //  Jobcode to filter the records
    "Chennai" // Location to filter the records
    );


3. If you have any queries, drop a mail to santhosh@tillidsoft.com. Thanks.
 