## Create workbook failure
In windows development debugging is correct. 
Deploying to docker reports an error: Create workbook failure
The solution is to install ttf support: apk add ttf-dejavu.

This is the alpine image file.