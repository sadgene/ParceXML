using Intermech;
using Intermech.Interfaces;
using Intermech.Interfaces.Services;
using Intermech.Interfaces.Server;
using Intermech.Interfaces.Client;
using Intermech.Interfaces.Compositions;
using Intermech.Interfaces.Workflow;
using Intermech.Workflow;
using Intermech.Expert.Scenarios;
using Intermech.Interfaces.Document;
using Intermech.Project;
using Intermech.Kernel;
using Intermech.Kernel.Search;
using System.ComponentModel.Design;
using Intermech.Interfaces.Projects;
using Intermech.Project;

using System.Collections.Generic;
using System;
using System.Windows.Forms;
using System.IO;
using System.Xml;
using System.Linq;
using System.Data;
using System.Text;
using System.ComponentModel.Design;


public class Script
{
    public ICSharpScriptContext ScriptContext { get; set; }
    public IUserSession session;
    StringBuilder strBuilder = new StringBuilder();
    
    public void Execute(IUserSession session, IServiceContainer services)
    {
        
        this.session = session;
        
        DirectoryInfo dir = new DirectoryInfo(@"C:\_Export1C");
        
        ReadXml(dir, session);
        
        
    }
    public void ReadXml(DirectoryInfo dir, IUserSession session)
    {
        IRouterService router = session.GetCustomService(typeof(IRouterService)) as IRouterService;
        string errorLog = String.Empty;
        
        string filenames = string.Empty;
        
        int filescount = 0;
        
        
        DateTime nowTime = DateTime.Now;
        
        
        strBuilder.AppendLine(DateTime.Now.ToString() + " Import task started");
        
        DirectoryInfo dirarchive = new DirectoryInfo(@"C:\_Export1C\XMLArchive");
        
        if (!dirarchive.Exists)
        {
            dirarchive.CreateSubdirectory("XMLArchive");
            strBuilder.AppendLine(DateTime.Now.ToString() + " XMLArchive directory created");
        }
        
        try
        {
            foreach (var item in dir.GetFiles())
            {
                
                
                if (item.Extension != ".xml")
                    continue;
                filescount++;
                strBuilder.AppendLine(DateTime.Now.ToString() + " Starting Parce XML : " + item.Name);//+Это
                
                filenames = dir.FullName + @"\" + item.Name;
                XmlDocument xDoc = new XmlDocument();
                xDoc.Load(filenames);
                
                XmlElement xRoot = xDoc.DocumentElement;
                
                foreach (XmlNode rowEl in xRoot.GetElementsByTagName("row"))
                {
                    
                    int numatr = 0;
                    
                    long userid = 0;
                    
                    List<int> numatributes = new List<int>();
                    foreach (XmlNode at in rowEl)
                    {
                        if (at.InnerText != "Null" & at.InnerText != String.Empty & at.InnerText != "")
                        {
                            numatributes.Add(numatr);
                        }
                        numatr++;
                        
                    }
                    
                    IDBObjectCollection colProObj = session.GetObjectCollection(1033);
                    IDBObject proobj = colProObj.Create();
                    bool iscreated = false;//был ли уже создан заказ
                    bool isgm = false;//признак гарантийное письмо
                    string subString = "технол";
                    int indexOfSubstring = -1;
                    foreach (XmlNode el1 in rowEl)
                    {
                        
                        foreach (XmlNode el in el1)
                        {
                            numatributes.Min();
                            
                            switch (numatributes.Min())
                            {
                                case 0:
                                
                                IDBObjectCollection colCodeObj = session.GetObjectCollection(1899);
                                
                                
                                foreach (DataRow row in colCodeObj.GetAttributeValues(14343, true).Rows)
                                {
                                    if (row[4].ToString() == el.Value)
                                    {
                                        iscreated = true;
                                    }
                                }
                                
                                if (!iscreated)
                                {
                                    IDBObject codeObj = colCodeObj.Create();
                                    codeObj.GetAttributeByName("Номер Заказа").AsString = el.Value;
                                    codeObj.CommitCreation(true);
                                    proobj.Attributes.AddAttribute(18122, false).Value = codeObj.ObjectID;
                                }
                                
                                IDBAttribute title = proobj.GetAttributeByName("Обозначение");
                                string s = el.Value;
                                
                                if (s.Length > 250)
                                {
                                    s = s.Substring(0, s.Length - (s.Length - 250));
                                }
                                
                                title.Value = s;
                                break;
                                case 1:
                                
                                break;
                                case 2:
                               
                                proobj.GetAttributeByName("Заказчик").AsString = el.Value;
                                break;
                                case 3:
                                
                                IDBAttribute proj_owner = proobj.Attributes.AddAttribute(18043, false);
                                
                                if (el.Value == "Отдел 1") proj_owner.Value = 1;
                                else if (el.Value == "Отдел 2") proj_owner.Value = 2;
                                else if (el.Value == "Отдел 3") proj_owner.Value = 3;
                                else if (el.Value == "Отдел 4") proj_owner.Value = 5;
                                else if (el.Value == "Отдел 5") proj_owner.Value = 7;
                                else if (el.Value == "Отдел 6") proj_owner.Value = 8;
                                else if (el.Value == "Отдел 7") proj_owner.Value = 9;
                                
                                break;
                                case 4:
                            
                                userid = this.GetUser(session, el.InnerText);
                                
                                if (userid == 0)
                                {
                                    userid = 126352161;
                                }
                                
                             
                                break;
                                case 5:
                               
                                if(el.InnerText == "ГП")
                                    isgm = true;
                                break;
                                case 6:
                               
                                break;
                                case 7:
                                proobj.GetAttributeByName("Дата окончания договора").AsString = el.Value.ToString();
                                break;
                                case 8:
                                
                                break;
                                case 9:
                                
                                break;
                                case 10:
                          
                                IDBAttribute proj_eosdo = proobj.Attributes.AddAttribute(18473, false);
                                proj_eosdo.AsString = el.InnerText;
                                break;
                                case 11:
                                
                                proobj.GetAttributeByName("Тип работ").AsString = el.Value.ToString();
                                break;
                                
                                case 12:
                                
                                string otdel = el.Value;
                               
                                indexOfSubstring = otdel.IndexOf(subString);
                             
                                
                                break;
                                case 13:
                                if (iscreated)
                                {
                                   
                                    strBuilder.AppendLine(DateTime.Now.ToString() + " Contract : " + proobj.GetAttributeByName("Обозначение").AsString + " Failed! Attribute " + "''Номер заказа''" + " Already Created ");
                                    break;
                                }
                                
                                IDBAttribute designatio = proobj.GetAttributeByName("Наименование");
                                string f = el.Value;
                                
                                if (f.Length > 250)
                                {
                                    f = f.Substring(0, f.Length - (f.Length - 250));
                                }
                                
                                designatio.Value = f;
                                
                                IDBAttribute template = proobj.GetAttributeByName("Шаблон проекта");
                                if(isgm)
                                {
                                    template.Value = 1234235123;
                                }
                                else
                                {
                                    template.Value = 191235235;
                                }
                          
                                
                                proobj.OwnerID = userid;
                                proobj.CommitCreation(true);
                                
                                proobj.OwnerID = 126352161;
                                
                                IDBProjectObject po = session.GetObject(proobj.ObjectID) as IDBProjectObject;
                                if(isgm)
                                {
                                    po.AddTemplateObjects(1272207);
                                }
                                else
                                {
                                    po.AddTemplateObjects(1946023);
                                }
                                
                                
                                proobj.Attributes.AddAttribute(18190, false);
                                proobj.GetAttributeByName("Ответсвенный исполнитель проекта (договора)").Value = userid;
                                if(indexOfSubstring >= 0 || isgm)
                                {
                                    
                                    proobj.GetAttributeByName("Готов к выдаче ").AsBoolean = true;
                                    proobj.GetAttributeByName("Готов к выдаче  текст").Value = "Заказ открыт";//
                                  
                                    
                                    router.CreateMessage(session.SessionGUID,userid/*userID*/,"Данные по договору "+proobj.GetAttributeByName("заказ").AsString+" перенесены в систему IPS","Договор : "+proobj.GetAttributeByName("Обозначение").AsString +
                                    "\n - Назначьте участников проекта \nЗаказ открыт автоматически.",2);
                                    
                                }
                                else
                                {
                              
                                    
                                    router.CreateMessage(session.SessionGUID,userid/*userID*/,"Данные по договору "+proobj.GetAttributeByName("заказ").AsString+" перенесены в систему IPS","Договор : "+proobj.GetAttributeByName("Обозначение").AsString +
                                    "\n - Назначьте участников проекта",2);
                                }
                                proobj.OwnerID = userid;
                                
                                
                                strBuilder.AppendLine(DateTime.Now.ToString() + " Contract : " + proobj.GetAttributeByName("Обозначение").AsString + " Success Created ");//+Это
                                
                                break;
                                case 15:
                               
                                proobj.Attributes.AddAttribute(18472, false);
                                proobj.GetAttributeByName("Гиперссылка на 1С").Value = el.InnerText;
                                break;
                            }
                            numatributes.Remove(numatributes.Min());
                        }
                    }
                }
                string new_path = Path.Combine(dirarchive.ToString(), Path.GetFileName(item.Name));
                strBuilder.AppendLine(DateTime.Now.ToString() + " Stopping Parce XML : " + item.Name);
                File.Move(item.FullName, new_path);
                strBuilder.AppendLine(DateTime.Now.ToString() + " " + item.Name + " - Removed to Archive");
            }
        }
        catch (Exception e)
        {
            errorLog = e.Message + "\n" + e.StackTrace;
        }
        
        if (errorLog != String.Empty)
        {
            File.AppendAllText(dirarchive.ToString() + "\\" + "importxml_from_1C" + DateTime.Now.Ticks.ToString() + ".error",
            errorLog, Encoding.UTF8);
            strBuilder.AppendLine(DateTime.Now.ToString() + " Error on export data " + errorLog);
        }
        
        if (filescount == 0)
        {
            strBuilder.AppendLine(DateTime.Now.ToString() + " No file to import was found");
            strBuilder.AppendLine(DateTime.Now.ToString() + " Import task ended");
            File.AppendAllText(dirarchive.ToString() + "\\importxml_from_1C.log", strBuilder.ToString(), Encoding.UTF8);
        }
        else
        {
            strBuilder.AppendLine(DateTime.Now.ToString() + " Import task ended");
            File.AppendAllText(dirarchive.ToString() + "\\importxml_from_1C.log", strBuilder.ToString(), Encoding.UTF8);
        }
    }
    public long GetUser(IUserSession session, string ownername)
    {
        IDBObjectCollection objColl = session.GetObjectCollection(new Guid("cad00002-306c-11d8-b4e9-00304f19f545" /*Пользователи*/));
        
        ConditionStructure[] conditions = new ConditionStructure[]
        {
            new ConditionStructure("Выводимое имя", RelationalOperators.Equal, ownername, LogicalOperators.NONE, 0, true)
        };
        
        ColumnDescriptor[] columnDescriptors = new ColumnDescriptor[]
        {
            new ColumnDescriptor(((int)ObligatoryObjectAttributes.F_OBJECT_ID), AttributeSourceTypes.Object, ColumnContents.ID,
            ColumnNameMapping.Default, SortOrders.NONE, -1),
        };
        
        // задаём параметры запроса
        DBRecordSetParams paramSet = new DBRecordSetParams(conditions, columnDescriptors, 0, null, QueryConsts.All);
        DataTable dtUsers = objColl.Select(paramSet);
        
        long idowner = 0;
        
        foreach (DataRow row in dtUsers.Rows)
        {
            idowner = Convert.ToInt64(row[0]);
        }
        return idowner;
    }
    
    
   
}
