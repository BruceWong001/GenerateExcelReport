using System;
using System.Data;
using System.IO;
using System.Reflection;
using System.Collections.Generic;


namespace GenerateExcelLib
{
    ///
    /// define for merge identifier.
    ///
    public class MergeIdentifierAttribute:Attribute
    {
        ///
        /// this Name will be set in rule dictionary, and used by merge follower obj.
        ///
        public string Name {get;set;}
        public Boolean IsHidden{get;set;}

        public MergeIdentifierAttribute(string identifierName,Boolean ishidden=false)
        {
            Name=identifierName;
            IsHidden=ishidden;
        }

    }
    ///
    /// define for merge follower
    ///
    public class MergeFollowerAttribute:Attribute
    {
        ///
        /// this Name will be set in rule dictionary, and used by merge follower obj.
        ///
        public string IdentifierName {get;set;}

        public MergeFollowerAttribute(string identifierName)
        {
            this.IdentifierName=identifierName;
        }

    }
    /// reflection can know current what is current data type.
    enum StructType
    {
        BasicType=0,
        GenericList,
        ComplexType

    } 


}