        ��  ��                  h      �� ��     0 	        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
   <assembly
     xmlns="urn:schemas-microsoft-com:asm.v1"
     manifestVersion="1.0">
     <assemblyIdentity
       processorArchitecture="x86"
       version="5.1.0.0"
       type="win32"   
       name="StartupList"/>
       <description>StartupList</description>
         <dependency>
           <dependentAssembly>
             <assemblyIdentity
               type="win32"
               name="Microsoft.Windows.Common-Controls"
               version="6.0.0.0"
               publicKeyToken="6595b64144ccf1df"
               language="*"
               processorArchitecture="x86"/>
           </dependentAssembly>
         </dependency>

<!-- Identify the application security requirements, for Windows Vista . -->
  <trustInfo xmlns="urn:schemas-microsoft-com:asm.v3">
     <security>
       <requestedPrivileges>
         <!-- <requestedExecutionLevel level="requireAdministrator" /> -->
           <requestedExecutionLevel level="highestAvailable"/>   
       </requestedPrivileges>
     </security>
  </trustInfo>

</assembly>
