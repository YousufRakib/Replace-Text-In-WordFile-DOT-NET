using Problem2.Models;
using SautinSoft.Document;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Web;
using System.Web.Mvc;

namespace Problem2.Controllers
{
    public class UserController : Controller
    {
        TestEntities db = new TestEntities();

        // GET: User
        public ActionResult UserList()
        {
            List<User> userInfos = new List<User>();
            userInfos=db.User.ToList();
            return View("UserList", userInfos);
        }
        public ActionResult ReplaceWordFile()
        {

            //string filePath = @"C:\Users\jakas\OneDrive\Desktop\Problem2\Problem2\Problem2\App_Data\File.docx";
            //DocumentCore dc = DocumentCore.Load(filePath);

            //// Searching string.
            //string searchText = "Rakib";

            //// Specify Regex to search with ignore case pattern:
            //// "TOM", "ToM", "toM", etc.
            //Regex regex = new Regex("(?i)" + searchText);
            //int count = dc.Content.Find(regex).Count();

            //var result = "The text: \"" + searchText + "\" - was found " + count.ToString() + " time(s).";

            return View();
        }
        public ActionResult FindText(HttpPostedFileBase upload, string existingText, string replaceText)
        {
            if (upload.ContentLength > 1)
            {
                string fileName = Path.GetFileName(upload.FileName);
                string fullFilePath = Path.Combine(Server.MapPath("~/App_Data"), fileName);

                if (System.IO.File.Exists(fullFilePath))
                {
                    var message = "This file already exist the system. Please change the file name";
                    ViewBag.Message = message.ToString();
                    return View("ReplaceWordFile");
                }
                else
                {
                    upload.SaveAs(fullFilePath);

                    DocumentCore dc = DocumentCore.Load(fullFilePath);

                    int countDel = 0;
                    Regex regex = new Regex("(?i)" + existingText.ToString());
                    foreach (ContentRange cr in dc.Content.Find(regex).Reverse())
                    {
                        cr.Replace(replaceText.ToString());
                        countDel++;
                    }

                    var _ext = Path.GetExtension(upload.FileName);

                    var FN = Path.GetFileNameWithoutExtension(fullFilePath);

                    string _comPath = Path.Combine(Server.MapPath("~/App_Data"), FN);
                    var path = _comPath +"_replace" + _ext;
                    
                    dc.Save(path);

                    var message = "File saved in "+ _comPath;
                    ViewBag.Message = message.ToString();
                }
            }
            return View("ReplaceWordFile");
        }
    }
}