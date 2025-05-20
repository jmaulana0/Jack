using Newtonsoft.Json.Linq;
using System.Reflection;

namespace DataUtilities
{
    public class DelegateInfo
    {
        public MethodInfo Action { get; }
        public object Instance { get; }
        public string Description { get; }
        public string ParametersSchema { get; }

        public string? ActionDescription { get; }
        public string? TargetTab { get; }
        public bool RequiresThreadID { get; set; }
        public bool RequiresWebSocket { get; set; }

        public DelegateInfo(MethodInfo action, object instance, string description, string parametersSchema, string actionDescription, string? targetTab = null, bool requiresThreadID = false, bool requiresWebSocket = false)
        {
            Action = action;
            Instance = instance;
            Description = description;
            ParametersSchema = parametersSchema;
            ActionDescription = actionDescription;
            TargetTab = targetTab;
            RequiresThreadID = requiresThreadID;
            RequiresWebSocket = requiresWebSocket;
        }
    }

    public static class DelegateInfoGenerator
    {
        public static List<DelegateInfo> GenerateDelegateInfos(object instance)
        {
            List<DelegateInfo> delegateInfos = new List<DelegateInfo>();
            MethodInfo[] methods = instance.GetType().GetMethods(BindingFlags.Instance | BindingFlags.NonPublic | BindingFlags.Public);

            foreach (MethodInfo method in methods)
            {
                DelegateInfoAttribute? attribute = method.GetCustomAttribute<DelegateInfoAttribute>();
                if (attribute != null)
                {
                    string schema = GenerateJsonSchemaAndContracts(method, attribute);

                    DelegateInfo delegateInfo = new DelegateInfo(
                        method,
                        instance,
                        attribute.Description,
                        schema,
                        attribute.ActionDescription,
                        attribute.TargetTab,
                        attribute.RequiresThreadID,
                        attribute.RequiresWebSocket
                    );
                    delegateInfos.Add(delegateInfo);
                }
            }

            return delegateInfos;
        }

        private static string GenerateJsonSchemaAndContracts(MethodInfo method, DelegateInfoAttribute? attribute)
        {
            ParameterInfo[] parameters = method.GetParameters();
            JObject properties = new JObject();
            JArray required = new JArray();

            foreach (ParameterInfo param in parameters)
            {
                if (param.GetCustomAttribute<ExcludeFromSchemaAttribute>() != null)
                {
                    continue;
                }

                JObject propertyObject = GeneratePropertySchema(param.ParameterType, param);

                properties[param.Name] = propertyObject;

                bool isRequired = !param.IsOptional;
                if (isRequired)
                {
                    required.Add(param.Name);
                }
            }

            JObject schema = new JObject
            {
                ["type"] = "object",
                ["properties"] = properties
            };

            if (required.Count > 0)
            {
                schema["required"] = required;
            }

            return schema.ToString();
        }

        private static string GetParameterDescription(ParameterInfo param)
        {
            var descAttr = param.GetCustomAttribute<ParameterDescriptionAttribute>();
            return descAttr?.Description ?? "";
        }

        private static JObject GeneratePropertySchema(Type type, ParameterInfo? parentParam = null)
        {
            // Handle nullable types  
            if (IsNullable(type, out Type underlyingType))
            {
                type = underlyingType;
            }

            // Handle array or list types  
            if (IsArrayOrListType(type))
            {
                Type elementType = GetElementType(type);
                return new JObject
                {
                    ["type"] = "array",
                    ["items"] = GeneratePropertySchema(elementType)
                };
            }

            // Handle enums  
            if (type.IsEnum)
            {
                var enumNames = Enum.GetNames(type);
                return new JObject
                {
                    ["type"] = "string",
                    ["enum"] = new JArray(enumNames)
                };
            }

            // Handle simple types  
            if (IsSimpleType(type))
            {
                return new JObject
                {
                    ["type"] = GetJsonType(type)
                };
            }

            // Handle complex objects  
            return GenerateObjectSchema(type);
        }

        private static JObject GenerateObjectSchema(Type type)
        {
            JObject properties = new JObject();
            JArray required = new JArray();

            var props = type.GetProperties(BindingFlags.Public | BindingFlags.Instance);
            foreach (var prop in props)
            {
                // Skip properties without a setter  
                if (!prop.CanWrite)
                    continue;

                JObject propertySchema = GeneratePropertySchema(prop.PropertyType);

                // Add property description if available  
                var descAttr = prop.GetCustomAttribute<PropertyDescriptionAttribute>();
                if (descAttr != null)
                {
                    propertySchema["description"] = descAttr.Description;
                }

                // Add property example if available  
                var exampleAttr = prop.GetCustomAttribute<PropertyExampleAttribute>();
                if (exampleAttr != null)
                {
                    propertySchema["example"] = JToken.FromObject(exampleAttr.Example);
                }

                properties[prop.Name] = propertySchema;

                // Assume properties are required unless they are nullable  
                if (!IsNullable(prop.PropertyType, out _))
                {
                    required.Add(prop.Name);
                }
            }

            JObject objSchema = new JObject
            {
                ["type"] = "object",
                ["properties"] = properties
            };

            if (required.Count > 0)
            {
                objSchema["required"] = required;
            }

            return objSchema;
        }

        private static bool IsNullable(Type type, out Type underlyingType)
        {
            if (type.IsGenericType && type.GetGenericTypeDefinition() == typeof(Nullable<>))
            {
                underlyingType = Nullable.GetUnderlyingType(type);
                return true;
            }
            underlyingType = type;
            return false;
        }

        private static bool IsSimpleType(Type type)
        {
            return
                type.IsPrimitive ||
                new Type[]
                {
                    typeof(string),
                    typeof(decimal),
                    typeof(DateTime),
                    typeof(Guid),
                    typeof(DateTimeOffset),
                    typeof(TimeSpan)
                }.Contains(type) ||
                type.IsEnum;
        }

        private static bool IsArrayOrListType(Type type)
        {
            return type.IsArray ||
                   (type.IsGenericType &&
                    (type.GetGenericTypeDefinition() == typeof(List<>) ||
                     type.GetGenericTypeDefinition() == typeof(IList<>) ||
                     type.GetGenericTypeDefinition() == typeof(IEnumerable<>)));
        }

        private static Type GetElementType(Type type)
        {
            if (type.IsArray)
                return type.GetElementType();

            if (type.IsGenericType)
                return type.GetGenericArguments()[0];

            return type;
        }

        private static string GetJsonType(Type type)
        {
            if (type == typeof(string))
                return "string";
            if (type == typeof(int) || type == typeof(long) || type == typeof(float) || type == typeof(double) || type == typeof(decimal))
                return "number";
            if (type == typeof(bool))
                return "boolean";
            if (type == typeof(DateTime) || type == typeof(DateTimeOffset))
                return "string";
            if (type == typeof(Guid))
                return "string";
            if (type == typeof(TimeSpan))
                return "string";
            return "object";
        }
    }
}