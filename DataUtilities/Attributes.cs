using System;

namespace DataUtilities
{
    [AttributeUsage(AttributeTargets.Property, Inherited = false, AllowMultiple = false)]
    public class PropertyDescriptionAttribute : Attribute
    {
        public string Description { get; }

        public PropertyDescriptionAttribute(string description)
        {
            Description = description;
        }
    }

    [AttributeUsage(AttributeTargets.Property, Inherited = false, AllowMultiple = false)]
    public class PropertyExampleAttribute : Attribute
    {
        public object Example { get; }

        public PropertyExampleAttribute(object example)
        {
            Example = example;
        }
    }

    [AttributeUsage(AttributeTargets.Method, Inherited = false, AllowMultiple = false)]
    public class DelegateInfoAttribute : Attribute
    {
        public string Description { get; }
        public string? ActionDescription { get; }
        public string? TargetTab { get; }
        public bool RequiresThreadID { get; }
        public bool RequiresWebSocket { get; }

        public DelegateInfoAttribute(string description, string? actionDescription = null, string? targetTab = null, bool requiresThreadID = false, bool requiresWebSocket = false)
        {
            Description = description;
            ActionDescription = actionDescription;
            TargetTab = targetTab;
            RequiresThreadID = requiresThreadID;
            RequiresWebSocket = requiresWebSocket;
        }
    }

    [AttributeUsage(AttributeTargets.Parameter, Inherited = false, AllowMultiple = false)]
    public class ParameterDescriptionAttribute : Attribute
    {
        public string Description { get; }

        public ParameterDescriptionAttribute(string description)
        {
            Description = description;
        }
    }

    [AttributeUsage(AttributeTargets.Parameter, Inherited = false, AllowMultiple = false)]
    public class ExcludeFromSchemaAttribute : Attribute
    {
    }
}