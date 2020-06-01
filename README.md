# Service Installer Test

Quick demo to show that the handle we get back from the Win32 CreateService API doesn't seem to care about the default ACL on the service and just gives us full access to the new service. Not a general security concern*, just an interesting thing

\* Normally only admins can create new services. Anyone you allow to create services could already create a service running as local system set to auto start then reboot the system even if they couldn't use this "trick" to allow themselves to start the service straight away.
