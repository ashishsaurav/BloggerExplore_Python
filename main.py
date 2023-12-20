from __future__ import print_function
import os
import sys
from pathlib import Path
from googleapiclient import sample_tools
import time
from xlwt import Workbook
import xlrd


class Post:
    def __init__(self):
        self.service = ""
        self.flags = ""
        self.posturl = []

    def signin(self, argv):
        try:
            if os.path.isfile("blogger.dat"):
                os.remove("blogger.dat")
            self.service, self.flags = sample_tools.init(argv, 'blogger', 'v3', __doc__, __file__,
                                                         scope='https://www.googleapis.com/auth/blogger')

        except:
            print("Sign in Unsuccessful")

    def createposts(self, blogid, title, content):
        try:
            posts = self.service.posts()
            body = {
                "kind": "blogger#post",
                "id": blogid,
                "title": title,
                "content": content
            }
            returnposturl = posts.insert(blogId=blogid, body=body, isDraft=False).execute()

            self.posturl.append(returnposturl['url'])

            return "Posted Successfully"
        except:
            return "Couldn't Post"

    def getpostsurl(self):
        try:
            if os.path.isfile("Url.xls"):
                os.remove("Url.xls")

            wb = Workbook()
            sheet1 = wb.add_sheet('Sheet1')
            rowno = 0
            for lst in self.posturl:
                sheet1.write(rowno, 0, lst)
                rowno += 1
            wb.save('Url.xls')
            print("Url fetched Successful")
        except:
            print("Url couldn't fetched")

    def getblodid(self):
        try:
            blogid = []
            # blogurl = []
            blogs = self.service.blogs()
            noofblogs = blogs.listByUser(userId='self').execute()

            for id in noofblogs['items']:
                blogid.append(id['id'])
                # blogurl.append(id['url'])

            if os.path.isfile("BloggerAPI.xls"):
                os.remove("BloggerAPI.xls")

            wb = Workbook()
            sheet1 = wb.add_sheet('Sheet1')
            rowno = 0
            for lst in blogid:
                # lst = "https://www.blogger.com/blog/settings/" + lst
                sheet1.write(rowno, 0, lst)
                rowno += 1

            # rowno = 0
            # for lst in blogurl:
            #     lst = lst[7:]
            #     lst = lst.split('.', 1)[0]
            #     sheet1.write(rowno, 1, lst)
            #     rowno += 1

            wb.save('BloggerAPI.xls')
            print("Id fetched Successful")

        except:
            print("Id couldn't fetch")

    def insertrecentposturl(self, blogid, title):
        try:
            posts = self.service.posts()
            allposts = posts.search(blogId=blogid, q=title).execute()
            self.posturl.append(allposts['items'][0]['url'])
            return "Recent Url fetched"
        except:
            return "Recent Url couldn't fetched"


if __name__ == '__main__':
    search = 'Invigolux Skin Serum Reviews - Anti-Aging Serum Solution!'
    post = Post()
    post.signin(sys.argv)
    post.getblodid()

    workbook = xlrd.open_workbook('BloggerAPI.xls')
    worksheet = workbook.sheet_by_name('Sheet1')
    txt = Path('Content.txt').read_text()
    title = Path('Title.txt').read_text()

    num_rows = worksheet.nrows
    curr_row = 0
    while curr_row < num_rows:
        blogid = worksheet.cell(curr_row, 0)
        er = post.insertrecentposturl(blogid.value, search)
        # er = post.createposts(blogid.value, title, txt)
        # if er != "Posted Successfully":
        #     print(str(curr_row+1) + ' ' + er)
        #     time.sleep(25)
        #     continue
        print(str(curr_row+1) + ' ' + er)
        curr_row += 1
    post.getpostsurl()
