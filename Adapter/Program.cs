
using Adapter;

UsersApi userRepository = new UsersApi();

IUserRepository adapter = new UsersRepositoryAdapter(userRepository);

List<string> users = adapter.GetUserNames();
foreach (var userName in users)
{
    Console.WriteLine(userName);
}
Console.WriteLine("----------------------------------");
Console.WriteLine();
//CSV 

IUserRepository csvAdapter = new UserReposytoryCsvAdapter(userRepository);
List<string> csvUsers = csvAdapter.GetUserNames();
foreach (var userName in csvUsers)
{
    Console.WriteLine(userName);
}