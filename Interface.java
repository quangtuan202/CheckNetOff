interface Interface {

    public String hello = "Hello";

    public void sayHello();
}

class InterfaceImpl implements Interface {

    public void sayHello() {
        System.out.println(Interface.hello);
    }
    public static void main(String args[]) {
        Interface x=new InterfaceImpl();
        x.sayHello();
    }
}

