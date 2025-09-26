import pandas as pd


data = {
'Enrollment Number': ['2203051240127', '2203051240088', '2203051240124', '2403051240022', '2203051240125', '2203051240126', '2203051240128', '2203051240129', '2203051240130', '2203051240131'],
'Name': ['John Doe', 'Priya Shah', 'Arjun Patel', 'Emma Wilson', 'Michael Chen', 'Sophia Rodriguez', 'Raj Kumar', 'Aisha Khan', 'David Smith', 'Mei Lin'],
'Email ID': ['john@example.com', 'priya@example.com', 'arjun@example.com', 'emma@example.com', 'michael@example.com', 'sophia@example.com', 'raj@example.com', 'aisha@example.com', 'david@example.com', 'mei@example.com'],
'Contact Number': ['9876543210', '8765432109', '7654321098', '6543210987', '5432109876', '4321098765', '3210987654', '2109876543', '1098765432', '0987654321'],
'DOB': ['01/01/2000', '15/05/2001', '22/07/1999', '30/11/2000', '05/03/2001', '18/09/2000', '27/02/1999', '12/12/2001', '08/06/2000', '19/04/2001'],
'Program': ['B.Tech', 'B.Sc', 'B.Tech', 'B.Com', 'B.Tech', 'B.A.', 'M.Tech', 'B.Sc', 'BBA', 'B.Arch'],
'Course': ['Computer Science', 'Physics', 'Artificial Intelligence', 'Commerce', 'Electronics', 'Psychology', 'Computer Science', 'Chemistry', 'Business Administration', 'Architecture'],
'Specialisation': ['Software Engineering', 'Quantum Physics', 'Machine Learning', 'Finance', 'VLSI Design', 'Clinical Psychology', 'Data Science', 'Organic Chemistry', 'Marketing', 'Urban Design'],
'Father Name': ['Robert Doe', 'Vikram Shah', 'Rajesh Patel', 'George Wilson', 'Wei Chen', 'Carlos Rodriguez', 'Suresh Kumar', 'Farhan Khan', 'James Smith', 'Zhang Lin'],
'Father Occupation': ['Businessman', 'Businessman', 'Businessman', 'Businessman', 'Engineer', 'Doctor', 'Professor', 'Lawyer', 'Accountant', 'Architect'],
'Profile Photo': ['https://randomuser.me/api/portraits/men/1.jpg', 'https://randomuser.me/api/portraits/women/2.jpg', 'https://randomuser.me/api/portraits/men/3.jpg', 'https://randomuser.me/api/portraits/women/4.jpg', 'https://randomuser.me/api/portraits/men/5.jpg', 'https://randomuser.me/api/portraits/women/6.jpg', 'https://randomuser.me/api/portraits/men/7.jpg', 'https://randomuser.me/api/portraits/women/8.jpg', 'https://randomuser.me/api/portraits/men/9.jpg', 'https://randomuser.me/api/portraits/women/10.jpg']
}


df = pd.DataFrame(data)
df.to_excel('students.xlsx', index=False)
print('Created students.xlsx with sample records.')