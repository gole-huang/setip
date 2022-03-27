# Renew IP Address
Setting a PC's IP Address within the information in a local Excel file
or in a MySQL DB Server. The program will find the information by the way 
of comparing the first column and the PC's current IP address. With the 
Excel file, it just renew the IP address , netmask and so on. With the 
MySQL DB Server, it renew the IP and report the result to the server.

### Prerequisites

1 All library file , you can download them from the "net452" floder
or get them with nuget.
2 .Net framework 4.5.2 runtime
3 MySQL DB Server is needed if you want to know how many stations have
been updated and how many hasn't.

## Running the tests

Explain how to run the automated tests for this system

### Sample Tests

Table in the IP.xlsx or MySQL has such column :
OLD_IP (string/varchar15), for current's IP address
NEW_IP (string/varchar15), for newest IP address
NEW_MASK (string/varchar15), for newest subnet mask
NEW_GATEWAY (string/varchar15), for newest gateway
NEW_DNS (string/varchar15), for newest DNS
IS_RENEW (tinybin), for whether has been renewed(MySQL only)
RENEW_DATE(UTC_DATETIME), for Renew's date(MySQL only)

## Built With
 -Dotnet 4.5.2 SDK
 -NPOI
 -MySqlConnector

## Authors

Gole Huang

## License

This project is licensed under the [MIT](LICENSE.md)

## Acknowledgments

  - Hat tip to anyone whose code is used
  - Inspiration
  - etc
