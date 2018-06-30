// <copyright file="GetRandomGUID.cs" >
// Copyright (c) 2017 All Rights Reserved
// </copyright>
// <author>Gokul JB</author>
// <date>6/19/2018 </date>
// <summary>Implements the GetRandomGUID Workflow Activity.</summary>
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Workflow;
using System;
using System.Activities;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Crm.Workflow.Utility
{
   public sealed class GetRandomGUID: CodeActivity
   {

       #region Pbulic Variable

       [Output("Random GUID")]
       public OutArgument<string> randomGuid { get; set; }

       #endregion

       protected override void Execute(CodeActivityContext executioncontext)
       {
           //It will generate random GUID and return to workflow
           randomGuid.Set(executioncontext, Convert.ToString(Guid.NewGuid()));
       }
    }
}
