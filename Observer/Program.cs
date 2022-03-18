
using Observer;

var newsAgency = new NewsAgency();

var dailyEconomy = new DailyEconomy();
var newYork = new NewYorkTimes();
var nationalGeographic = new NationalGeographic();

newsAgency.Attach(dailyEconomy);
newsAgency.Attach(newYork);
newsAgency.Attach(nationalGeographic);


newsAgency.SetNewsHeadline(Genre.Economy, "USA is going bancrupt!");
//newsAgency.Notify();

newsAgency.SetNewsHeadline(Genre.Science, "Life on Alpha Centauri");
//newsAgency.Notify();

newsAgency.SetNewsHeadline(Genre.Sport, "Adam Małysz is the greatest sportsman in the history of mankind");
//newsAgency.Notify();

newsAgency.SetNewsHeadline(Genre.Economy, "CD Project RED value has grown by 500% in 2020");
//newsAgency.Notify();

newsAgency.SetNewsHeadline(Genre.Science, "Kirkendall effect causing airplanes' engine deteriorate");
//newsAgency.Notify();

newsAgency.Detach(dailyEconomy);

newsAgency.SetNewsHeadline(Genre.Economy, "Texas is going bancrupt!");
//newsAgency.Notify();


newsAgency.Detach(newYork);
newsAgency.Detach(nationalGeographic);