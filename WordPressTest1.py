from wordpress_xmlrpc import Client, WordPressPost;
from wordpress_xmlrpc.methods.posts import GetPosts, NewPost;
from wordpress_xmlrpc.methods.users import GetUserInfo;

def newPost():
    print("newPost");
    wp = Client('http://39.106.104.45/wordpress/xmlrpc.php', 'shikun', 'ShiKun001')
    post = WordPressPost();
    post.title = 'Lichuan Test5';
    post.content = 'This is the body of Lichuan Test5.';
    post.terms_names = {
    'post_tag': ['test', 'lichuan'],
    'category': ['Introductions', 'Tests']
    };
    post.id = wp.call(NewPost(post));
    print("post.id = " + str(post.id) );

if __name__ == '__main__' :
    print("begin");
    newPost();
    print("end.");

