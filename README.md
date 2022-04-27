gridMail
===================
__Command line based emailing program that used Microsoft Exchange Services__


## Requirements

* .NET Framework v4.0+
* Exchange Account

## Build Requirements
* Visual Studio 2015
* [.Net Framework v4.0](http://www.microsoft.com/en-us/download/details.aspx?id=17851)
* [Exchange Web Services (EWS) API](https://www.microsoft.com/en-us/download/details.aspx?id=42022)


## Usage
Simple:
 ```
 gridMail.exe -t ""person1@domain.com"" -b ""Body goes here"" -s ""Subject Line"" 
 ```


Attachments:
```
gridMail.exe -a ""item1.txt,item2.txt"" -t ""person1@domain.com,person2@email.com"" -b ""Body goes here"" -s ""Subject Line""
```


HTML:
```
gridMail.exe -t ""person1@domain.com"" -w ""body_template.html"" -s ""Subject Line""
```


HTML w/images:
```
gridMail.exe -t ""person1@domain.com" -s ""Subject"" -w ""index.html"" -i "picture.jpg@123456;path\to\picture.jpg"
```

