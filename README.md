
# Excel CRUD API

Excel CRUD App is a Spring Boot application that performs CRUD operations on an excel file (.xlsx) through RESTful APIs with the help of Apache POI library and used Postman for API testing

Developed using Java, Spring Boot, Spring Web, Apache POI.


## Installation

### Using Git ###
1. Clone the project from github.

```bash
  git clone https://github.com/nitish-nagar/ExcelCrudApp.git .
```

2. Create an excel file in your C drive of your local machine and you can name db.xlsx and your path should be as such:

C:\db.xlsx

3. Open the project in Eclipse IDE and right click on project and select "Run as Java Application"

### APIs ###

getUserList
```bash
  GET: http:localhost:8080/users
```

getUserById
```bash
  GET: http:localhost:8080/users/{id}
```

createUser
```bash
  POST: http:localhost:8080/users
```

addUserList
```bash
  POST: http:localhost:8080/userList
```

updateUser
```bash
  PUT: http:localhost:8080/users/{id}
```

modifyUserList
```bash
  PUT: http:localhost:8080/userList
```

deleteUser
```bash
  DELETE: http:localhost:8080/users/{id}
```
## Authors

- [Nitish Nagar](https://github.com/nitish-nagar)