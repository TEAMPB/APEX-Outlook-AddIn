# APEX-Outlook-AddIn
Tool zum Erstellen von APEX-Outlook-AddIns

## API

https://learn.microsoft.com/de-de/javascript/api/outlook/office.mailbox?view=outlook-js-preview

[Nur Add-In-Manifestreferenz für Office-Add-Ins - Office Add-ins | Microsoft Learn](https://learn.microsoft.com/de-de/javascript/api/manifest?view=outlook-js-preview)





## App

* no authentication App, public-Page oder Office365 Authentication (zur Not)
* App-Definition > Security: Embed in Frames **Allow**
* App-Definition > User Interface oder Page : JavaScript / File URLs : https://appsforoffice.microsoft.com/lib/1.1/hosted/office.js





## Page

Function and Global Variable Declaration

```javascript


Office.onReady((info)=>{
    console.log('Info:')
    console.log(info);

    console.log('Office Context:')
    console.log(Office.context);


    apex.item('P1_USER').setValue(Office.context.mailbox.userProfile.emailAddress);
    apex.item('P1_SENDER').setValue(Office.context.mailbox.item.from.emailAddress);
    
    Office.context.mailbox.item.body.getAsync(Office.CoercionType.Html,(result)=>{
	      apex.item('P1_BODY').setValue(result.value);
    });

});
```



## In Outlook installieren

Installieren: https://aka.ms/olksideload
