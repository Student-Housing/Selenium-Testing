Imports System.Text
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports OpenQA.Selenium
Imports OpenQA.Selenium.Firefox
Imports OpenQA.Selenium.Chrome
Imports OpenQA.Selenium.Support.UI

Namespace SeleniumTesting

    <TestClass()> Public Class TestGoogleSearching
        Private Const GECKO_FOLDER As String = "C:\Users\Scott K\Desktop\FirefoxGeckoDriver\"
        Private Const GOOGLE As String = "http://www.google.com/"
        Private Firefox As IWebDriver

        <TestInitialize()>
        Public Sub Initialization()
            Firefox = New FirefoxDriver(GECKO_FOLDER)
        End Sub
    
        <TestCleanup()>
        Public Sub Termination()
            Firefox.Quit()
        End Sub
    
        <TestMethod()> Public Sub GoogleSearch()
            Firefox.Navigate.GoToUrl(GOOGLE)

            With Firefox.FindElement(By.Name("q"))
                .SendKeys(PickSomeWords())
                .SendKeys(Keys.Enter)
            End With

            System.Threading.Thread.Sleep(2000)
        End Sub

        <TestMethod()> Public Sub GoogleSearchAndListLinks()
            GoogleSearch()

            Dim SearchResult As IWebElement

            For Each SearchResult In Firefox.FindElements(By.PartialLinkText(".com"))
                System.Diagnostics.Debug.WriteLine(SearchResult.Text)
            Next SearchResult

            System.Threading.Thread.Sleep(2000)
        End Sub

        <TestMethod()> Public Sub GoogleSearchAndClickFirstLink()
            GoogleSearch()
            Firefox.FindElement(By.PartialLinkText(".com")).Click()

            System.Threading.Thread.Sleep(2000)
        End Sub

        Private Function PickSomeWords() As String
            Dim SearchTerms As Collection = New Collection
            Dim SearchTermCount As Integer

            With SearchTerms
                .Add("Why")
                .Add("What")
                .Add("How")
                .Add("Where")
                .Add("Time")
                .Add("A")
                .Add("Ultimate")
                .Add("Best")
                .Add("Remember")
                .Add("BnB")
                .Add("Canada")
                .Add("Flag")
                .Add("America")
                .Add("Cookies")
                .Add("Chocolate-Chip")
                .Add("Roomba")
                .Add("FFIX")
                .Add("Penguin")
                .Add("+ 1")
                .Add("1")
                .Add("2")
                .Add("3")
                .Add("Guitar")
                .Add("Piano")
                .Add("Cello")
            End With

            SearchTermCount = SearchTerms.Count

            Dim I As Integer

            For I = 2 To RandomBetween(5, 2)
                PickSomeWords = PickSomeWords & SearchTerms.Item(RandomBetween(SearchTermCount)) & " "
            Next I
        End Function

        Private Function RandomBetween(ByVal UpperBound As Integer, Optional ByVal LowerBound As Integer = 1) As Integer
            Randomize()
            RandomBetween = Int((UpperBound - LowerBound + 1) * Rnd() + LowerBound)
        End Function
    End Class

End Namespace
