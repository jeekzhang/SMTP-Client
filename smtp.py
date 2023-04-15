import tkinter as tk
from tkinter import messagebox, ttk
from tkinter import END
from socket import *
import ssl
from email.base64mime import body_encode
import re
import time

mailServer = "smtp.qq.com"
# 发送方地址和接收方地址，from 和 to
#填入发送方和收件方
fromAddress = "1458836984@qq.com"
toAddress = "jeekzhang@139.com"
# 发送方，验证信息，由于邮箱输入信息会使用base64编码，因此需要进行编码
username = "1458836984@qq.com"  # 输入自己的用户名对应的编码
password = ""  # 此处不是自己的密码，而是开启SMTP服务时对应的授权码
endMsg = "\r\n.\r\n"
contentType = "text/plain"
pattern = '{"发件人": "(.*?)", "收件人": "(.*?)", "主题": "(.*?)", "内容": "(.*?)", "发送时间": "(.*?)"}\n'
patch = re.compile(pattern)
pattern_a = '{"发件人": "(.*?)", "收件人": "(.*?)", "主题": "(.*?)", "内容": "(.*?)", "保存时间": "(.*?)"}\n'
patch_a = re.compile(pattern_a)
pattern_b = '{"联系人": "(.*?)", "邮箱": "(.*?)"}\n'
patch_b = re.compile(pattern_b)


class App:

    def __init__(self, master):
        self.notebook = ttk.Notebook(master)
        self.frame1 = tk.Frame(master)
        self.frame2 = tk.Frame(master)
        self.frame3 = tk.Frame(master)
        self.frame4 = tk.Frame(master)

        # frame1
        self.notebook.add(self.frame1, text='写邮件')
        self.notebook.add(self.frame2, text='通讯录')
        self.notebook.add(self.frame3, text='已发送')
        self.notebook.add(self.frame4, text='草稿箱')
        self.notebook.pack(padx=10, pady=5, fill=tk.BOTH, expand=True)
        self.server = tk.StringVar()
        self.server.set(mailServer)
        self.sender = tk.StringVar()
        self.sender.set(fromAddress)
        self.receiver = tk.StringVar()
        self.receiver.set(toAddress)
        self.username = tk.StringVar()
        self.username.set(username)
        self.password = tk.StringVar()
        self.password.set(password)
        self.title = tk.StringVar()
        ttk.Label(self.frame1, text='SMTP服务器:').grid(
            row=0, column=0, sticky='w', pady=5)
        ttk.Entry(self.frame1, textvariable=self.server,
                  state='readonly').grid(row=0, column=1, sticky='w')
        ttk.Label(self.frame1, text='发送人:').grid(
            row=1, column=0, sticky='w', pady=5)
        ttk.Entry(self.frame1, textvariable=self.sender,
                  state='readonly').grid(row=1, column=1, sticky='w')
        ttk.Label(self.frame1, text='收件人:').grid(
            row=2, column=0, sticky='w', pady=5)
        ttk.Entry(self.frame1, textvariable=self.receiver,
                  width=50).grid(row=2, column=1, sticky='w')
        ttk.Label(self.frame1, text='主题:').grid(
            row=5, column=0, sticky='w', pady=5)
        ttk.Entry(self.frame1, textvariable=self.title).grid(
            row=5, column=1, sticky='w')
        ttk.Label(self.frame1, text='邮件内容:').grid(
            row=6, column=0, sticky='w', pady=5)
        ybar = ttk.Scrollbar(self.frame1, orient='vertical')
        self.textarea = tk.Text(self.frame1, width=50,
                                height=10, yscrollcommand=ybar.set)
        ybar['command'] = self.textarea.yview
        self.textarea.grid(row=6, column=1, columnspan=1, sticky='w')
        ybar.grid(row=6, column=2, sticky='ns')
        ttk.Button(self.frame1, text='点击发送', command=lambda: self.sending(
            self.textarea.get(1.0, 'end'))).grid(row=10, column=2, sticky='w')
        ttk.Button(self.frame1, text='存至草稿', command=lambda: self.draft(
            self.textarea.get(1.0, 'end'))).grid(row=10, column=1, sticky='w')
        ttk.Button(self.frame1, text='清空内容', command=lambda: self.clear(
            self.textarea)).grid(row=10, column=0, sticky='w')

        # frame3
        self.scrollbar = tk.Scrollbar(self.frame3)
        self.scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        title = ['1', '2', '3', '4', '5', ]
        self.box = ttk.Treeview(self.frame3, columns=title,
                                yscrollcommand=self.scrollbar.set,
                                show='headings')
        self.box.column('1', width=170, anchor='center')
        self.box.column('2', width=170, anchor='center')
        self.box.column('3', width=60, anchor='center')
        self.box.column('4', width=190, anchor='center')
        self.box.column('5', width=170, anchor='center')

        self.box.heading('1', text='发件人')
        self.box.heading('2', text='收件人')
        self.box.heading('3', text='主题')
        self.box.heading('4', text='内容')
        self.box.heading('5', text='发送时间')
        # 对象处理
        self.refresh()

        # frame4
        self.scrollbar_a = tk.Scrollbar(self.frame4)
        self.scrollbar_a.pack(side=tk.RIGHT, fill=tk.Y)

        self.box_a = ttk.Treeview(self.frame4, columns=title,
                                  yscrollcommand=self.scrollbar_a.set,
                                  show='headings')
        self.box_a.bind('<Double-1>', self.treeviewClick_a)
        self.box_a.column('1', width=170, anchor='center')
        self.box_a.column('2', width=170, anchor='center')
        self.box_a.column('3', width=60, anchor='center')
        self.box_a.column('4', width=190, anchor='center')
        self.box_a.column('5', width=170, anchor='center')

        self.box_a.heading('1', text='发件人')
        self.box_a.heading('2', text='收件人')
        self.box_a.heading('3', text='主题')
        self.box_a.heading('4', text='内容')
        self.box_a.heading('5', text='保存时间')
        # 对象处理
        self.refresh_a()

        # frame2
        ttk.Button(self.frame2, text='添加联系人',
                   command=lambda: self.addcontact()).pack(side=tk.TOP)
        self.scrollbar_b = tk.Scrollbar(self.frame2)
        self.scrollbar_b.pack(side=tk.RIGHT, fill=tk.Y)
        actitle = ['1', '2', ]
        self.box_b = ttk.Treeview(self.frame2, columns=actitle,
                                  yscrollcommand=self.scrollbar_b.set,
                                  show='headings')
        self.box_b.bind('<Double-1>', self.treeviewClick_b)
        self.box_b.column('1', width=350, anchor='center')
        self.box_b.column('2', width=400, anchor='center')

        self.box_b.heading('1', text='收件人名字')
        self.box_b.heading('2', text='收件人邮箱')
        # 对象处理
        self.refresh_b()

    def treeviewClick_b(self, event):

        for item in self.box_b.selection():
            item_text = self.box_b.item(item, "values")
            self.receiver.set(item_text[1])

    def treeviewClick_a(self, event):
        for item in self.box_a.selection():
            item_text = self.box_a.item(item, "values")
            self.receiver.set(item_text[1])
            self.title.set(item_text[2])
            self.clear(self.textarea)
            self.textarea.insert("end", item_text[3].replace("\\n", "\n"))

    def refresh(self,):
        obj = self.box.get_children()  # 获取所有对象
        for o in obj:
            self.box.delete(o)
        self.op = self.readdata('sended.txt')
        while 1:
            try:
                line = next(self.op)
            except StopIteration:
                break
            else:
                result = patch.match(line)
                if (result != None):
                    self.box.insert('', 'end', values=[
                                    result[i] for i in range(1, 6)])

        self.scrollbar.config(command=self.box.yview)
        self.box.pack(side=tk.LEFT, fill=tk.Y)

    def refresh_a(self,):
        obj = self.box_a.get_children()  # 获取所有对象
        for o in obj:
            self.box_a.delete(o)
        self.op_a = self.readdata('draft.txt')
        while 1:
            try:
                line = next(self.op_a)
            except StopIteration:
                break
            else:
                result = patch_a.match(line)
                if (result != None):
                    self.box_a.insert('', 'end', values=[
                                      result[i] for i in range(1, 6)])

        self.scrollbar_a.config(command=self.box_a.yview)
        self.box_a.pack(side=tk.LEFT, fill=tk.Y)

    def refresh_b(self,):
        obj = self.box_b.get_children()  # 获取所有对象
        for o in obj:
            self.box_b.delete(o)
        self.op_b = self.readdata('contact.txt')
        while 1:
            try:
                line = next(self.op_b)
            except StopIteration:
                break
            else:
                result = patch_b.match(line)
                if (result != None):
                    self.box_b.insert('', 'end', values=[
                                      result[i] for i in range(1, 3)])

        self.scrollbar_b.config(command=self.box_b.yview)
        self.box_b.pack(side=tk.LEFT, fill=tk.Y)

    def readdata(self, path):
        """逐行读取文件"""

        f = open(path, 'r', encoding='utf-8')
        line = f.readline()
        while line:
            yield line
            line = f.readline()

        f.close()

    def sending(self, content):  # 发送邮件
        sender = self.sender.get()
        receivers = self.receiver.get().split()  # 多个收件人

        # 创建客户端套接字并建立连接
        serverPort = 465  # SMTP使用587号端口，测试SSL使用465号端口
        clientSocket = socket(AF_INET, SOCK_STREAM)
        # SSl加密，默认版本为TSL
        clientSocket = ssl.wrap_socket(
            clientSocket, keyfile='./privkey.pem', certfile='./certificate.pem', server_side=False)
        clientSocket.connect((mailServer, serverPort))  # connect只能接收一个参数
        # 从客户套接字中接收信息
        recv = clientSocket.recv(1024).decode()
        # print(recv)
        if '220' != recv[:3]:
            messagebox.showerror('错误', '220 reply not received from server.')

        # 发送 HELO 命令并且打印服务端回复
        # 开始与服务器的交互，服务器将返回状态码250,说明请求动作正确完成
        heloCommand = 'HELO Alice\r\n'
        clientSocket.send(heloCommand.encode())  # 随时注意对信息编码和解码
        recv1 = clientSocket.recv(1024).decode()
        # print(recv1)
        if '250' != recv1[:3]:
            messagebox.showerror('错误', '250 reply not received from server.')

        # 发送"AUTH PLAIN"命令，验证身份.服务器将返回状态码334（服务器等待用户输入验证信息）
        user_pass_encode64 = body_encode(
            f"\0{username}\0{password}".encode('ascii'), eol='')
        clientSocket.sendall(f'AUTH PLAIN {user_pass_encode64}\r\n'.encode())
        clientSocket.recv(1024).decode()
        # print(recv2)

        # 发送 MAIL FROM 命令，并包含发件人邮箱地址
        clientSocket.sendall(('MAIL FROM: <' + fromAddress + '>\r\n').encode())
        recvFrom = clientSocket.recv(1024).decode()
        # print(recvFrom)
        if '250' != recvFrom[:3]:
            messagebox.showerror('错误', '250 reply not received from server')

        # 发送 RCPT TO 命令，并包含收件人邮箱地址，返回状态码 250
        clientSocket.sendall(('RCPT TO: <' + toAddress + '>\r\n').encode())
        recvTo = clientSocket.recv(1024).decode()  # 注意UDP使用sendto，recvfrom
        # print(recvTo)
        if '250' != recvTo[:3]:
            messagebox.showerror('错误', '250 reply not received from server')

        # 发送 DATA 命令，表示即将发送邮件内容。服务器将返回状态码354（开始邮件输入，以"."结束）
        clientSocket.send('DATA\r\n'.encode())
        recvData = clientSocket.recv(1024).decode()
        # print(recvData)
        if '354' != recvData[:3]:
            messagebox.showerror('错误', '354 reply not received from server')

        # 编辑邮件信息，发送数据

        message = 'from:' + sender + '\r\n'
        message += 'to:' + receivers[0] + '\r\n'
        message += 'subject:' + self.title.get() + '\r\n'
        message += 'Content-Type:' + contentType + '\t\n'
        message += '\r\n' + content
        clientSocket.sendall(message.encode())

        # 以"."结束。请求成功返回 250
        clientSocket.sendall(endMsg.encode())
        recvEnd = clientSocket.recv(1024).decode()
        # print(recvEnd)
        if '250' != recvEnd[:3]:
            print('错误', '250 reply not received from server')

        # 发送"QUIT"命令，断开和邮件服务器的连接
        clientSocket.sendall('QUIT\r\n'.encode())

        clientSocket.close()

        with open('sended.txt', "a", encoding='gbk') as filewrite:  # ”a"代表着每次运行都追加txt的内容
            s_sended = "{\"发件人\": \""+sender+"\", \"收件人\": \""+receivers[0]+"\", \"主题\": \""+self.title.get(
            )+"\", \"内容\": \""+content.replace("\n", "\\n")+"\", \"发送时间\": \""+str(time.asctime())+"\"}\n"
            filewrite.write(s_sended)
        self.refresh()
        messagebox.showinfo('提示', '邮件发送完成')

    def clear(self, textarea):  # 清空文本框
        textarea.delete(1.0, 'end')

    def draft(self, content):
        sender = self.sender.get()
        receivers = self.receiver.get().split()
        with open('draft.txt', "a", encoding='gbk') as filewrite:  # ”a"代表着每次运行都追加txt的内容
            s_draft = "{\"发件人\": \""+sender+"\", \"收件人\": \""+receivers[0]+"\", \"主题\": \""+self.title.get(
            )+"\", \"内容\": \""+content.replace("\n", "\\n")+"\", \"保存时间\": \""+str(time.asctime())+"\"}\n"
            filewrite.write(s_draft)
        self.refresh_a()
        messagebox.showinfo('提示', '草稿保存完成')

    def addcontact(self,):
        def addDate():
            n = name.get()
            p = ph.get()
            if n == '' or p == '':
                messagebox.showerror("错误", "信息有不能为空!")
                win1.destroy()
            else:
                with open('contact.txt', "a", encoding='gbk') as filewrite:  # ”a"代表着每次运行都追加txt的内容
                    s_contact = "{\"联系人\": \""+n+"\", \"邮箱\": \""+p+"\"}\n"
                    filewrite.write(s_contact)
                self.refresh_b()
                messagebox.showinfo("成功", "新增联系人成功")
                win1.destroy()
        win1 = tk.Toplevel()
        win1.title('新增联系人')
        win1.geometry('500x300')
        sw = win1.winfo_screenwidth()
        sh = win1.winfo_screenheight()
        win1.geometry('+%d+%d' % ((sw - 500) / 2, (sh - 300) / 2))
        # 欢迎语
        l = tk.Label(win1, text='新增联系人', font=('华文行楷', 20), fg='purple')
        l.place(relx=0.5, rely=0.1, anchor='center')
        # 提示语
        l = tk.Label(win1, text='请输入联系人信息', font=('正楷', 15))
        l.place(relx=0.5, rely=0.3, anchor='center')
        # 姓名输入框
        Lname = tk.Label(win1, text='姓名:')
        Lname.place(relx=0.1, rely=0.5, anchor='center')
        nu = tk.StringVar()
        name = tk.Entry(win1, textvariable=nu)
        name.place(relx=0.15, rely=0.47, width=70)
        # 邮箱输入框
        Lph = tk.Label(win1, text='邮箱:')
        Lph.place(relx=0.4, rely=0.5, anchor='center')
        nu1 = tk.StringVar()
        ph = tk.Entry(win1, textvariable=nu1)
        ph.place(relx=0.45, rely=0.47, width=140)
        # 按钮
        b = tk.Button(win1, text='添加', bg="gray", command=addDate)
        b.place(relx=0.75, rely=0.45, width=100)


if __name__ == '__main__':
    root = tk.Tk()
    root.title("SMTP邮件客户端")
    w_width = 800  # 工具宽度
    w_height = 400  # 工具高度
    scn_width = root.maxsize()[0]  # 屏幕宽度
    x_point = (scn_width-w_width)//2  # 取点让工具居中
    root.geometry('%dx%d+%d+%d' % (w_width, w_height, x_point, 100))
    App(root)
    root.mainloop()
