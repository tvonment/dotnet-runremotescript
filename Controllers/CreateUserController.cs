using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using System;
using System.IO;
using System.Threading.Tasks;
using System.Collections.Generic;
using Microsoft.AspNetCore.Http;
using Newtonsoft.Json;
using System.Management.Automation;
using System.Management.Automation.Runspaces;
using System.Security;
using RunRemotePowershell.Models;
using System.Reflection;

namespace RunRemotePowershell.Controllers
{
    [ApiController]
    [Route("api/v1/[controller]")]
    public class CreateUserController : ControllerBase
    {
        private readonly ILogger<CreateUserController> _logger;

        public CreateUserController(ILogger<CreateUserController> logger)
        {
            _logger = logger;
        }

        [HttpPost]
        public string Post([FromBody] User user)
        {
            var scriptParams = new Dictionary<string, string>();
            PropertyInfo[] properties = typeof(User).GetProperties();
            foreach (PropertyInfo property in properties)
            {
                if (property.GetValue(user) != null)
                {
                    scriptParams.Add(property.Name, property.GetValue(user).ToString());
                } else
                {
                    scriptParams.Add(property.Name, "");
                }
            }

            return RunRemoteScript(scriptParams);
        }

        private string RunRemoteScript(Dictionary<string, string> scriptParams)
        {
            WSManConnectionInfo connectionInfo = new WSManConnectionInfo();
            connectionInfo.Credential = new PSCredential(Environment.GetEnvironmentVariable("User"), ConvertToSecureString(Environment.GetEnvironmentVariable("Password")));
            connectionInfo.ComputerName = Environment.GetEnvironmentVariable("ScriptHost");
            Runspace runspace = RunspaceFactory.CreateRunspace(connectionInfo);
            runspace.Open();

            string outputString = "";

            using (PowerShell ps = PowerShell.Create())
            {
                ps.Runspace = runspace;

                string scriptPath = Environment.GetEnvironmentVariable("ScriptPath");
                // specify the script code to run.
                ps.AddCommand(scriptPath);

                // specify the parameters to pass into the script.
                ps.AddParameters(scriptParams);

                // execute the script and await the result.                
                var output = ps.Invoke();

                // print the resulting pipeline objects to the console.


                foreach (var item in output)
                {
                    Console.WriteLine(item.ToString());
                    outputString += item.ToString() + "\n";
                }
            }
            runspace.Close();

            string responseMessage = string.IsNullOrEmpty(outputString)
                ? "NO OUTPUT"
                : $"Connect to {Environment.GetEnvironmentVariable("ScriptHost")} as {Environment.GetEnvironmentVariable("User")} with OUTPUT: \n\n{outputString}";

            return responseMessage;

            SecureString ConvertToSecureString(string password)
            {
                if (password == null)
                    throw new ArgumentNullException("password");

                var securePassword = new SecureString();

                foreach (char c in password)
                    securePassword.AppendChar(c);

                securePassword.MakeReadOnly();
                return securePassword;
            }
        }
    }
}
