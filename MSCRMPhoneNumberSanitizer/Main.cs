﻿using System;
using System.Net;
using System.ServiceModel.Description;
using Microsoft.Xrm.Sdk.Client;
using Xrm;
using System.Linq;
using System.Text.RegularExpressions;

namespace MSCRMPhoneNumberSanitizer
{
	class MainClass
	{
        const int BatchSize = 20;

		public static void Main (string[] args)
		{
            Console.Clear();

            if (args == null || args.Length < 3)
            {
                Console.WriteLine("usage is ");
                Console.WriteLine();
                Console.WriteLine("MSCRMPhoneNumberSanitizer.exe MSCRMURL USERNAME PASSWORD DOMAIN");
                Console.WriteLine();
                Console.WriteLine("MSCRMURL should be in the format http://jon-mscrm.dev2000.com/MSCRM/XRMServices/2011/Organization.svc");
                Console.WriteLine("User name and password are for a user that has rights to update records in MSCRM");
                Console.WriteLine();
            }
            else
            {
                try
                {
                    var credentials = new ClientCredentials();
                    credentials.Windows.ClientCredential = new NetworkCredential(args[1], args[2], args[3]);
                    var serviceProxy = GetServiceProxy(new Uri(args[0]), credentials);

                    UpdateContacts(serviceProxy);
                    UpdateAccounts(serviceProxy);
                }
                catch (Exception ex)
                {
                    Console.WriteLine("ERROR: " + ex.Message);
                }
            }

            Console.WriteLine();
            Console.WriteLine("Press any key to quit...");
            Console.Read();
		}

		private static OrganizationServiceProxy GetServiceProxy(Uri serviceUri, ClientCredentials credentials)
		{
            // Suppress prompt for user credentials
            credentials.SupportInteractive = false;

            var proxy = new OrganizationServiceProxy(serviceUri, null, credentials, null);

            proxy.ServiceConfiguration.CurrentServiceEndpoint.Behaviors.Add(new ProxyTypesBehavior());

            return proxy;
		}

        private static void UpdateContacts(OrganizationServiceProxy serviceProxy)
        {
            var context = new XrmServiceContext(serviceProxy);
            var updateCount = 0; 
            foreach (var contact in context.ContactSet)
            {
                Console.WriteLine(String.Format("{0} {1}|{2}|{3}|{4}", contact.FullName,
                contact.MobilePhone,
                contact.Telephone1,
                contact.Telephone2,
                contact.Telephone3));
       
                var newMobilePhone = SanitizeNumber(contact.MobilePhone);
                var newTelephone1 = SanitizeNumber(contact.Telephone1);
                var newTelephone2 = SanitizeNumber(contact.Telephone2);
                var newTelephone3 = SanitizeNumber(contact.Telephone3);

                if (newMobilePhone != contact.MobilePhone ||
                    newTelephone1 != contact.Telephone1 ||
                    newTelephone2 != contact.Telephone2 ||
                    newTelephone3 != contact.Telephone3)
                {
                    contact.MobilePhone = newMobilePhone;
                    contact.Telephone1 = newTelephone1;
                    contact.Telephone2 = newTelephone2;
                    contact.Telephone3 = newTelephone3;
                    context.UpdateObject(contact);
                    updateCount++;
                }

                if (updateCount == BatchSize)
                {
                    updateCount = 0;
                    context.SaveChanges();
                }
            }

            //final update
            if (updateCount > 0)
            {
                context.SaveChanges();
            }
        }

        private static void UpdateAccounts(OrganizationServiceProxy serviceProxy)
        {
            var context = new XrmServiceContext(serviceProxy);
            var updateCount = 0;
            foreach (var account in context.AccountSet)
            {
                Console.WriteLine(String.Format("{0} {1}|{2}|{3}", account.Name,
                account.Telephone1,
                account.Telephone2,
                account.Telephone3));

                var newTelephone1 = SanitizeNumber(account.Telephone1);
                var newTelephone2 = SanitizeNumber(account.Telephone2);
                var newTelephone3 = SanitizeNumber(account.Telephone3);

                if (newTelephone1 != account.Telephone1 ||
                    newTelephone2 != account.Telephone2 ||
                    newTelephone3 != account.Telephone3)
                {
                    account.Telephone1 = newTelephone1;
                    account.Telephone2 = newTelephone2;
                    account.Telephone3 = newTelephone3;
                    context.UpdateObject(account);
                    updateCount++;
                }

                if (updateCount == BatchSize)
                {
                    updateCount = 0;
                    context.SaveChanges();
                }
            }

            //final update
            if (updateCount > 0)
            {
                context.SaveChanges();
            }
        }

        private static string SanitizeNumber(string number)
        {
            if (String.IsNullOrEmpty(number))
            {
                return number;
            }
            return Regex.Replace(number, "\\D", String.Empty);
        }
	}
}
