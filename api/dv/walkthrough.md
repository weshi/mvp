_[Home page](../index.md)_


# Walkthrough
The following function will read all the properties from the document and a few custom properties into a task pane:

```javascript
    function getAllProperties() {
        Excel.run(function(context) {
            var docProperties = context.workbook.properties;
            var customDocProperties = context.workbook.properties.custom;
            var customProperty1 = customDocProperties.getItem("FirstValue");
            var customProperty2 = customDocProperties.getItem("SecondValue");
            var customProperty3 = customDocProperties.getItem("ThirdValue");
            // Load a combination of read-only 
            // and writeable document properties.
            docProperties.load("author, lastAuthor, revisionNumber, title, subject, keywords, comments, category, manager, company, creationDate");
            customProperty1.load("key, value");
            customProperty2.load("key, value");
            customProperty3.load("key, value");
            return context.sync().then( function() { 
                // Write the document properties to the console.
                // To learn how to view document properties in the UI, 
                // see https://support.office.com/en-us/article/View-or-change-the-properties-for-an-Office-file-21D604C2-481E-4379-8E54-1DD4622C6B75
                $("#doc-prop-author").val(docProperties.author);
                $("#doc-prop-last").val(docProperties.lastAuthor);
                $("#doc-prop-revision").val(docProperties.revisionNumber);
                $("#doc-prop-title").val(docProperties.title);
                $("#doc-prop-subject").val(docProperties.subject);
                $("#doc-prop-keywords").val(docProperties.keywords);
                $("#doc-prop-comments").val(docProperties.comments);
                $("#doc-prop-category").val(docProperties.category);
                $("#doc-prop-manager").val(docProperties.manager);
                $("#doc-prop-company").val(docProperties.company);
                $("#doc-prop-create").val(docProperties.creationDate.toDateString());
                // Write the custom document properties to the console.
                // To learn how to view document properties in the UI, 
                // see https://support.office.com/en-us/article/View-or-change-the-properties-for-an-Office-file-21D604C2-481E-4379-8E54-1DD4622C6B75
                $("#custom-doc-property-firstvalue").val(customProperty1.value);
                $("#custom-doc-property-secondvalue").val(customProperty2.value);
                $("#custom-doc-property-thirdvalue").val(customProperty3.value);
            });
        });
    }
```

The following function will write a subset of document properties and three custom document properties:

```javascript
    function setCustomDocProperties() {
        Excel.run(function(context) {
            var customDocProperties = context.workbook.properties.custom;
            var docProperties = context.workbook.properties;

            // update specific properties
            docProperties.title = $("#doc-prop-title").val();
            docProperties.subject = $("#doc-prop-subject").val();
            docProperties.keywords = $("#doc-prop-keywords").val();
            docProperties.comments = $("#doc-prop-comments").val();
            docProperties.category = $("#doc-prop-category").val();
            docProperties.manager = $("#doc-prop-manager").val();
            docProperties.company = $("#doc-prop-company").val();

            // Add custom document properties.
            customDocProperties.add("FirstValue", $("#custom-doc-property-firstvalue").val());
            customDocProperties.add("SecondValue", $("#custom-doc-property-secondvalue").val());
            customDocProperties.add("ThirdValue", $("#custom-doc-property-thirdvalue").val());

            return context.sync().then( function(context) {
                $("#message-banner").show();
                messageBanner.showBanner();
            });
        });
    }
```
