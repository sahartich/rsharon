----- Script by Sahar Tichover -----

Usage Instructions

Type RSHARON to search for a computer by name and connect to it.

If you already know the computer name, type RDPSHARON to connect directly.

🔄 Functionality Overview
Automatically updates CSV files with relevant data.

Queries local CSV files to locate target computers.

Checks which computers are online and accepting RDP (TCP port 3389) connections.

Prompts the user to select an available computer to connect to.

🛠️ Version History & Changes
v1.3.3

Refactored into multiple smaller files for improved readability, modularity, and easier maintenance.

v1.3.2

Active Directory data is now saved in an encrypted user folder as a CSV file and queried locally for faster access.

v1.3.1

Quality-of-life improvements:

Ability to search again at any point.

Clean exit without using Ctrl+C.

Improved error handling.

v1.3

Replaced Test-NetConnection with .NET's System.Net.Sockets.TcpClient for faster and more efficient port checking.

Implemented AsyncWaitHandle to timeout after 1 second, significantly reducing total runtime.

v1.2

On first run, the script creates a complete list of all users and computers from Active Directory, eliminating the need for repeated queries.

v1.1

Checks each computer to ensure it accepts TCP connections on the RDP port before listing it as available.