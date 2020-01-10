# AD-Computer-Create

Utility for quickly creating computer accounts in Active Directory.

## Features

- Create a single computer account in a specified OU
- Create multiple computer accounts in a specified OU from an input list
  - Can be imported from a semicolon delimited text file
  - Accounts created using a naming convention (changes would require editing)
- config.xml
  - Machine Prefix & Location Code can be specified to override script defaults.
  - AD realm/region/container can be specified to override defaults.
  - Created accounts can be added to group policy objects defined in config.xml
    - Separate policy groups may be specified for laptops & desktops

## Prerequisites

- Windows 10 RSAT (Remote Server Administration Tools) installed
  - 1809 and up: Optional Features - `RSAT: Active Directory Domain Services and Lightweight Directory Services Tools`

## How to use

Run `AD-Computer-Create.ps1` (or .exe if compiled). The configuration file `config.xml` needs to be located in the same folder as the executable else it will use built-in defaults.

### Single Computer Account

- Computer Name: The exact name of the computer account to create.
- User's Name: The user's name will be inserted into the Description field.
- OU: The container path will automatically be generated from the _Location_ and _Machine Type_ choices below.
  - Location: The 3-letter designation for the computer's location. e.g. DAY, RDU, HBE, etc.
  - Machine Type: Select either Mac, Laptop, or Desktop
- When the `Create Computer Account` button is clicked:
  - If the computer name already exists in Active Directory, a warning will written to the console window and execution will terminate.
  - The computer account is then created in the specified container as shown in the _OU_ field.
  - If set in config.xml, a specified group will have proper permissions set on the computer account to be able to join the computer to the domain, otherwise _Domain Admins_ is the default group for this.
  - If set in config.xml, the computer account is then added to specified group policy objects.

### Multiple Computer Accounts

- Input: A _semicolon delimited_ text file may be imported or lines may be manually entered.

  - Each line should contain a unique identifier (such as the asset tag or serial number) and a description (i.e. user's name).
  - Example:

    ```text
    7001234;Smith, John
    7005678;Doe, Jane
    ```

- OU: The container path will automatically be generated from the _Location_ and _Machine Type_ choices below.
  - Location: The 3-letter designation for the computer's location. e.g. DAY, RDU, HBE, etc.
  - Machine Type: Select either Laptop, or Desktop
- When the `Create Computer Accounts` button is clicked, the following will occur for each line:
  - The computer name will be generated as:
    - `[PREFIX][LOCATION][L or D]-[IDENTIFIER]`
    - e.g. `CMPDAYL-7001234`
    - The prefix can be changed in config.xml
    - - This naming convention was used because of my needs at the time of script creation. Fine-tuning to the code may be required to suit your own requirements.
  - If the computer name already exists in Active Directory, a warning will written to the console window and the line is skipped.
  - The computer account is then created in the specified container as shown in the _OU_ field.
  - If set in config.xml, a specified group will have proper permissions set on the computer account to be able to join the computer to the domain, otherwise _Domain Admins_ is the default group for this.
  - If set in config.xml, the computer account is then added to specified group policy objects.

### config.xml

Below is a sample configuration file.

```xml
<?xml version="1.0"?>
<ADComputerCreate>
  <!--
      Computer Naming Convention: <MachinePrefix><LocationCode><MachineType>-<ID>
      Example: CMPDAYL-7001234
  -->
  <MachinePrefix>CMP</MachinePrefix>
  <LocationCode>DAY</LocationCode>
  <!--
      OU defined as:
      OU=<LocationCode>,OU=<Region>,<Laptops|Desktops>,OU=<Container>,<Realm>

      Example: OU=DAY,OU=Americas,OU=Laptops,OU=Workstations,DC=domain,DC=forest,DC=tld
  -->
  <AD>
    <Realm>DC=domain,DC=forest,DC=tld</Realm>
    <Region>Americas</Region>
    <Container>Workstations</Container>
    <JoinDomainGroup>Client Installers</JoinDomainGroup>
  </AD>
  <!--
      Newly created computer accounts can be added to defined group policies
  -->
  <Laptops>
    <Policies>
      <Policy>Certificate Enabled Workstations</Policy>
      <Policy>MBAM BitLocker Encrypt</Policy>
    </Policies>
    <Attributes>
    </Attributes>
  </Laptops>
  <Desktops>
    <Policies>
    </Policies>
    <Attributes>
    </Attributes>
  </Desktops>
  <Macs>
    <Policies>
      <Policy>Certificate Enabled Workstations</Policy>
    </Policies>
    <Attributes>
      <!-- Value will be pulled from OU realm -->
      <!-- Example: cmpdayl-7001234.domain.forest.tld -->
      <Attribute>dnsHostName</Attribute>
    </Attributes>
  </Macs>
</ADComputerCreate>
```
