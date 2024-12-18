# Introduction 
TODO: Give a short introduction of your project. Let this section explain the objectives or the motivation behind this project. 

# Getting Started

submodule instructions:

1.	Navigate to the root directory of your solution: Open the integrated terminal in Visual Studio and navigate to the root directory of your solution where you want to add the submodule.
2.	Add the submodule: Use the git submodule add command to add the submodule. Replace <repository-url> with the URL of the repository you want to add as a submodule and <path> with the directory path where you want to place the submodule.

> git submodule add https://dacgrouplabs.visualstudio.com/DefaultCollection/Core%20Services/_git/Dac.ReportExporter Dac.ReportExporter

3.	Initialize and update the submodule: After adding the submodule, initialize and update it to fetch the content.

> git submodule update --init --recursive

4.	Commit the changes: Commit the changes to your repository to include the submodule.

> git add .gitmodules Dac.ReportExporter
> git commit -m "Add submodule https://dacgrouplabs.visualstudio.com/DefaultCollection/Core%20Services/_git/Dac.ReportExporter at Dac.ReportExporter"

5.	Include the submodule in your solution: In Visual Studio, you can now include the submodule in your solution. Right-click on your solution in the Solution Explorer, select "Add" -> "Existing Project...", and navigate to the submodule directory to add the project file.

6. **Handle Submodule Changes Carefully**
- **Committing Changes**: When making changes to a submodule, commit those changes in the submodule repository first, then update the main repository to point to the new commit.

cd submodule-directory
git add .
git commit -m "Update submodule"
cd ..
git add submodule-directory
git commit -m "Update submodule reference"


7. **Use Git Hooks for Automation**
- **Post-Checkout Hook**: Use Git hooks to automate submodule updates. For example, a post-checkout hook can automatically update submodules after checking out a branch.
    
# .git/hooks/post-checkout
#!/bin/sh
git submodule update --init --recursive

8. **Monitor Submodule Status**
- **Status Check**: Regularly check the status of submodules to ensure they are in sync.

> git submodule status


9. **Avoid Frequent Submodule Changes**
- **Stability**: Try to keep submodule references stable. Frequent changes can lead to merge conflicts and complicate the development process.

10. **CI/CD Integration**
- **Build Scripts**: Ensure your CI/CD pipelines are configured to initialize and update submodules.
    

# Build and Test
TODO: Describe and show how to build your code and run the tests. 

# Contribute
TODO: Explain how other users and developers can contribute to make your code better. 

If you want to learn more about creating good readme files then refer the following [guidelines](https://docs.microsoft.com/en-us/azure/devops/repos/git/create-a-readme?view=azure-devops). You can also seek inspiration from the below readme files:
- [ASP.NET Core](https://github.com/aspnet/Home)
- [Visual Studio Code](https://github.com/Microsoft/vscode)
- [Chakra Core](https://github.com/Microsoft/ChakraCore)