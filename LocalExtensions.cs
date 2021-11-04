using Microsoft.Graph;
using System.Reflection;
using System.Text;

namespace TeamsExplorer
{
    public static class LocalExtensions
    {
        public static string ToStringExtended(this TeamFunSettings settings)
        {
            return GetString(settings);
        }

        public static string ToStringExtended(this TeamGuestSettings settings)
        {
            return GetString(settings);
        }

        public static string ToStringExtended(this TeamMemberSettings settings)
        {
            return GetString(settings);
        }

        public static string ToStringExtended(this TeamMessagingSettings settings)
        {
            return GetString(settings);
        }

        private static string GetString<T>(T obj)
        {
            StringBuilder sb = new StringBuilder();
            sb.Append("{");
            int iteration = 0;
            var props = obj.GetType().GetProperties();
            foreach (PropertyInfo p in props)
            {
                iteration++;
                sb.Append(p.Name);
                sb.Append(":");
                string value = p.GetValue(obj, null) != null ? p.GetValue(obj, null).ToString() : "null";
                sb.Append(value);
                if(iteration<props.Length)
                    sb.Append(", ");
            }
            
            sb.Append("}");

            return sb.ToString();
        }
    }
}
