// =====================================================================
//  This file is part of the Microsoft Dynamics CRM SDK code samples.
//
//  Copyright (C) Microsoft Corporation.  All rights reserved.
//
//  This source code is intended only as a supplement to Microsoft
//  Development Tools and/or on-line documentation.  See these other
//  materials for detailed information regarding Microsoft code samples.
//
//  THIS CODE AND INFORMATION ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY
//  KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE
//  IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A
//  PARTICULAR PURPOSE.
// =====================================================================

//<snippetAutoRouteLead>
using System;
using System.Activities;
using System.Collections.ObjectModel;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Workflow;

namespace Futuredontics.ParentAccountUpdateonLead
{
    public sealed class ParentAccountUpdateonLead : CodeActivity
    {
        /// <summary>
        /// This method first retrieves the lead. Afterwards, it checks the Parent Account id
        /// If Existing Account has data , all the Campaign responses of Lead are rolled upto Account else we remove the Cr's from Account that are related to Lead
        /// </summary>


        protected override void Execute(CodeActivityContext executionContext)
        {

            #region Create the tracing service
            ITracingService tracingService = executionContext.GetExtension<ITracingService>();
            #endregion

            #region Create the context
            IWorkflowContext context = executionContext.GetExtension<IWorkflowContext>();
            IOrganizationServiceFactory serviceFactory = executionContext.GetExtension<IOrganizationServiceFactory>();
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            #endregion

            #region Retrieve the Contact GUID
            Guid contactid = context.PrimaryEntityId;
            tracingService.Trace("Lead Guid -" + contactid);
            #endregion
            Entity campaignresponse = new Entity("campaignresponse");
            EntityReference acc = new EntityReference("account");

            #region Get Parent Account Id using Query by Attribute
            QueryByAttribute AccountQueryBycontactId = new QueryByAttribute("contact");
            AccountQueryBycontactId.AddAttributeValue("contactid", contactid);
            AccountQueryBycontactId.ColumnSet = new ColumnSet("contactid", "parentcustomerid");
            EntityCollection contactrecords = service.RetrieveMultiple(AccountQueryBycontactId);


            foreach (Entity contactrecord in contactrecords.Entities)
            {
                Guid cr = Guid.Empty;
                string fetchContactCR = "<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>" +
                                                    "<entity name='campaignresponse'>" +
                                                      "<attribute name='subject' />" +
                                                      "<attribute name='activityid' />" +
                                                      "<attribute name='regardingobjectid' />" +
                                                      "<attribute name='responsecode' />" +
                                                      "<attribute name='customer' />" +
                                                      "<attribute name='statecode' />" +
                                                      "<attribute name='statuscode' />" +
                                                      "<order attribute='subject' descending='false' />" +
                                                      "<filter type='and'>" +
                                                        "<condition attribute='fdx_reconversioncontact' operator='eq'  value='" + contactid + "' />" +
                                                      "</filter>" +
                                                    "</entity>" +
                                                  "</fetch>";
                EntityCollection ContactCrs = service.RetrieveMultiple(new FetchExpression(fetchContactCR));
                if (ContactCrs.Entities.Count > 0)
                {
                    tracingService.Trace("ContactCR Count-" + ContactCrs.Entities.Count);

                    foreach (Entity contactcr in ContactCrs.Entities)
                    {
                        cr = ((Guid)contactcr["activityid"]);

                        tracingService.Trace("CR Guid-" + cr);

                        if (contactrecord.Attributes.Contains("parentcustomerid"))
                        {
                            acc.Id = ((EntityReference)contactrecord["parentcustomerid"]).Id;
                            acc.Name = ((EntityReference)contactrecord["parentcustomerid"]).Name;

                            tracingService.Trace("Account Guid-" + acc.Id);
                            tracingService.Trace("Account Name -" + acc.Name);

                            EntityReference ContactAccount = new EntityReference("account", acc.Id);

                            Entity customer = new Entity("activityparty");
                            customer.Attributes["partyid"] = ContactAccount;
                            EntityCollection Customerentity = new EntityCollection();
                            Customerentity.Entities.Add(customer);

                            campaignresponse["customer"] = Customerentity;

                            tracingService.Trace("Customer -" + Customerentity);

                            campaignresponse["activityid"] = cr;

                            service.Update(campaignresponse);

                            tracingService.Trace("CR is updated with Account on Customer from Contact");
                        }
                        else
                        {
                            campaignresponse["customer"] = null;

                            campaignresponse["activityid"] = cr;

                            service.Update(campaignresponse);

                            tracingService.Trace("Account is cleared on Customer-Campaign Response as Contact is untagged");
                        }
                    }
                }
            #endregion
            }
        }
    }
}
