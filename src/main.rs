use clap::{Arg, command};
use reqwest::blocking::Client;
use rust_xlsxwriter::{Workbook, XlsxError};
use scraper::{Html, Selector};
use std::error::Error;
use urlencoding::encode;

#[derive(Debug)]
struct Article {
    title: String,
    authors: String,
    link: String,
}

fn scrape(query: &str, num_pages: u32) -> Vec<Article> {
    let client = Client::builder()
        .user_agent("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36")
        .build()
        .unwrap();

    let mut articles = Vec::new();
    let result_selector = Selector::parse("div.gs_ri").unwrap();
    let title_selector = Selector::parse("h3.gs_rt").unwrap();
    let authors_selector = Selector::parse("div.gs_a").unwrap();
    let anchor_selector = Selector::parse("a").unwrap();

    for page in 0..num_pages {
        let url = format!(
            "https://scholar.google.com/scholar?start={}&q={}&hl=en&as_sdt=0,5",
            page * 10,
            encode(query)
        );

        if let Ok(response) = client.get(&url).send() {
            if let Ok(body) = response.text() {
                let document = Html::parse_document(&body);

                for element in document.select(&result_selector) {
                    let title = element
                        .select(&title_selector)
                        .next()
                        .map(|el| el.text().collect())
                        .unwrap_or_else(|| "No title found".into());

                    let authors = element
                        .select(&authors_selector)
                        .next()
                        .map(|el| el.text().collect())
                        .unwrap_or_else(|| "No authors info".into());

                    let link = element
                        .select(&anchor_selector)
                        .next()
                        .and_then(|el| el.value().attr("href"))
                        .unwrap_or("No link")
                        .to_string();

                    articles.push(Article {
                        title,
                        authors,
                        link,
                    });
                }
            }
        }
    }

    articles
}

fn save(articles: &[Article], filename: &str) -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();

    worksheet.write_string(0, 0, "Title")?;
    worksheet.write_string(0, 1, "Authors")?;
    worksheet.write_string(0, 2, "Link")?;

    for (i, article) in articles.iter().enumerate() {
        let row = (i + 1) as u32;
        worksheet.write_string(row, 0, &article.title)?;
        worksheet.write_string(row, 1, &article.authors)?;
        worksheet.write_string(row, 2, &article.link)?;
    }

    workbook.save(filename)
}

fn main() -> Result<(), Box<dyn Error>> {
    let mut query = "";
    let mut pages = 1;
    let filename;

    let res = command!()
        .author("Pranav V Bhat")
        .about("CLI tool to scrape articles from Google Scholar")
        .arg_required_else_help(true)
        .arg(
            Arg::new("query")
                .short('q')
                .long("query")
                .required(true)
                .help("Specify a string of keywords to search for"),
        )
        .arg(
            Arg::new("pages")
                .short('p')
                .long("pages")
                .required(false)
                .help("Number of pages to scrape (1 page = 10 articles; default: 1)"),
        )
        .arg(
            Arg::new("save")
                .short('s')
                .long("save")
                .required(false)
                .help("Path to the .xlsx file you want to save in (default: query name)"),
        )
        .get_matches();

    let arg_query = res.get_one::<String>("query");
    let arg_pages = res.get_one::<String>("pages");
    let arg_path = res.get_one::<String>("save");

    if let Some(arg_query) = arg_query {
        query = arg_query;
    }

    if let Some(arg_pages) = arg_pages {
        pages = arg_pages.parse::<u32>()?;
    }

    if let Some(arg_path) = arg_path {
        filename = arg_path.to_string();
    } else {
        filename = format!("{}.xlsx", query);
    }

    println!(
        "Scraping Google Scholar for '{}' ({} pages)...",
        query, pages
    );
    let articles = scrape(query, pages);

    println!("Found {} articles. Saving to Excel...", articles.len());
    save(&articles, &filename)?;

    println!("Successfully saved results to {}", filename);
    Ok(())
}
