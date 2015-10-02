using System;
using System.Collections.Generic;
#if tiger_dotNet4
using System.Linq;
#endif
using System.Text;
using Route2NSEx.src.Marquis.systemUtil;
using System.Reflection;
using System.Text.RegularExpressions;

namespace MarsUFTAddins.IMars.tiger
{
    public class ReflectorForCSharp
    {
        private static MLogger Logger = MLogger.GetLogger(typeof(ReflectorForCSharp));

        protected string getParameterAndValue(string strPara, string strV)
        {
            return string.Format("{0}:[{1}]", strPara, strV);
        }

        public T GetMember<T>(object objSrc, string strMemIdx)
        {
            Logger.logBegin("GetMember");
            Logger.Info("GetMember", getParameterAndValue("strMemIdx", strMemIdx));
            Type objType = objSrc.GetType();
            MemberInfo[] arrMember = objType.GetMember(strMemIdx, BindingFlags.Instance | BindingFlags.NonPublic | BindingFlags.Public);
            if (arrMember.Length <= 0)
            {
                Logger.Info("GetMember", "no such Member:" + strMemIdx);
                return default(T);
            }
            MemberInfo objOne = arrMember[0];
            if (objOne == null)
            {
                Logger.Info("GetMember", "find member but value is null:" + strMemIdx);
                return default(T);
            }
            switch (objOne.MemberType)
            {
                case MemberTypes.Field:
                    Logger.Info("GetMember", "return field");
                    T objResult = (T)(((FieldInfo)objOne).GetValue(objSrc));
                    Logger.logEnd("GetMember");
                    return objResult;
                case MemberTypes.Property:
                    Logger.Info("GetMember", "return Porperty");
                    objResult = (T)(((PropertyInfo)objOne).GetValue(objSrc, null));
                    Logger.logEnd("GetMember");
                    return objResult;
                default:
                    Logger.Info("GetMember", "unsupported: " + objOne.MemberType.ToString());
                    Logger.logEnd("GetMember");
                    return default(T);
            }

        }
        public T GetPrivateField<T>(object instance, string fieldname)
        {
            Logger.logBegin("GetPrivateField");
            Logger.Info("GetPrivateField", getParameterAndValue("fieldname", fieldname));
            BindingFlags flag = BindingFlags.Instance | BindingFlags.NonPublic;
            Type type = instance.GetType();
            FieldInfo field = type.GetField(fieldname, flag);
            Logger.logEnd("GetPrivateField");
            if (field == null) return default(T);
            return (T)field.GetValue(instance);
        }

        public void GetAllEventsName(object objInst, ref List<string> lstResult)
        {
            Logger.logBegin("--------GetAllEventsName--------");
            EventInfo[] arrList = objInst.GetType().GetEvents(BindingFlags.Public | BindingFlags.Instance | BindingFlags.NonPublic);
            if (lstResult == null) lstResult = new List<string>();
            foreach (EventInfo e in arrList)
            {
                lstResult.Add(string.Format("{0,32}", e.Name));
            }
            Logger.logEnd("GetAllEventsName");
        }

        public T GetPrivateProperty<T>(object instance, string propertyname)
        {
            Logger.logBegin("GetPrivateProperty");
            Logger.Info("GetPrivateProperty", getParameterAndValue("propertyname", propertyname));
            BindingFlags flag = BindingFlags.Instance | BindingFlags.NonPublic;
            Type type = instance.GetType();

            PropertyInfo field = type.GetProperty(propertyname, flag);
            Logger.logEnd("GetPrivateProperty");
            if (field == null) return default(T);
            return (T)field.GetValue(instance, null);
        }

        public void SetPrivateField(object instance, string fieldname, object value)
        {
            Logger.logBegin("SetPrivateField");
            Logger.Info("SetPrivateField", getParameterAndValue("fieldname", fieldname));
            Logger.Info("SetPrivateField", getParameterAndValue("value", value == null ? "" : value.ToString()));
            BindingFlags flag = BindingFlags.Instance | BindingFlags.NonPublic;
            Type type = instance.GetType();
            FieldInfo field = type.GetField(fieldname, flag);
            field.SetValue(instance, value);
            Logger.logEnd("SetPrivateField");
        }

        public void SetPrivateProperty(object instance, string propertyname, object value)
        {
            Logger.logBegin("SetPrivateProperty");
            Logger.Info("SetPrivateProperty", getParameterAndValue("propertyname", propertyname));
            Logger.Info("SetPrivateProperty", getParameterAndValue("value", value == null ? "" : value.ToString()));
            BindingFlags flag = BindingFlags.Instance | BindingFlags.NonPublic;
            Type type = instance.GetType();
            PropertyInfo field = type.GetProperty(propertyname, flag);
            field.SetValue(instance, value, null);
            Logger.logEnd("SetPrivateProperty");
        }

        public T CallPrivateMethod<T>(object instance, string name, params object[] param)
        {
            Logger.logBegin("CallPrivateMethod");
            Logger.Info("CallPrivateMethod", getParameterAndValue("propertyname", name));
            BindingFlags flag = BindingFlags.Instance | BindingFlags.NonPublic;
            Type type = instance.GetType();
            MethodInfo method = type.GetMethod(name, flag);
            if (method == null) return default(T);
            return (T)method.Invoke(instance, param);
        }

        public void CallMethod(object inst, string strName, params object[] param)
        {
            Logger.logBegin("CallMethod");
            Logger.Info("CallMethod", getParameterAndValue("propertyname", strName));
            BindingFlags flag = BindingFlags.Instance | BindingFlags.NonPublic | BindingFlags.Public;
            Type type = inst.GetType();
            MethodInfo method = type.GetMethod(strName, flag);
            method.Invoke(inst, param);
            return;
        }

        public void CallEvent(object objInst, string strEvent, params object[] param)
        {
            Logger.logBegin("CallEvent");
            EventInfo objEvnt = objInst.GetType().GetEvent(strEvent);
            if (objEvnt == null)
            {
                Logger.Info("CallEvent", string.Format("no such event:{0}", strEvent));
                return;
            }
            Type objDeleType = objEvnt.EventHandlerType;
            MethodInfo objMthdInfoHandle = objInst.GetType().GetMethod(strEvent, BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Static);
            Delegate d = Delegate.CreateDelegate(objDeleType, objMthdInfoHandle);
            //MethodInfo objAddHandler = objEvnt.get

            Logger.logEnd("CallEvent");
        }

        public MethodInfo GetMethod(object objInst, string strMethodName)
        {
            Logger.logBegin("GetMethod");
            try
            {
                if (objInst == null) return null;

                MethodInfo objMethod = objInst.GetType().GetMethod(strMethodName, BindingFlags.Instance | BindingFlags.NonPublic | BindingFlags.Public);
                return objMethod;
            }
            finally
            {
                Logger.logEnd("GetMethod");
            }
        }


        public sealed class MarsTigerUtility
        {
            public static MLogger Logger = MLogger.GetLogger(typeof(MarsTigerUtility));
            public static bool RegularExpressChecking(string strPartern, string strSrc)
            {
                Logger.logBegin(string.Format("RegularExpressChecking,{0}-{1}", CombinePara("strPartern", strPartern),CombinePara("strSrc", strSrc)));
                try 
	            {	        
		            Regex regex = new Regex(strPartern) ;
                    return regex.IsMatch(strSrc) ; 
	            }
	            finally
	            {
                    Logger.logEnd("RegularExpressChecking") ;
	            }                
            }

            
            public static string CombinePara(string strPara, string strValue)
            {
                return string.Format("{0}:[{1}]", strPara, strValue);
            }
        }

    }
}
