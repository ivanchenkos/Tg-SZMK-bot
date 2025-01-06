pub mod parse;
use calamine::{
    Data, 
    Range
};
use teloxide::{
    dispatching::{dialogue::{self, InMemStorage}, UpdateHandler},
    prelude::*,
    types::{InlineKeyboardButton, InlineKeyboardMarkup},
    utils::command::BotCommands,
};

type MyDialogue = Dialogue<State, InMemStorage<State>>;
type HandlerResult = Result<(), Box<dyn std::error::Error + Send + Sync>>;

#[derive(Clone, Default)]
pub enum State {
    #[default]
    Start,
    InlineChooseMode,
    InlineRecieveMode,
    CablesRecieveNumber, 
}

/// These commands are supported:
#[derive(BotCommands, Clone)]
#[command(rename_rule = "lowercase")]
enum Command {
    Start,
    Cancel,
} 

pub fn write_log(log: &str) {
    log::info!("{}", log);
}

pub fn write_log_in_bot(log: &str) {
    todo!()
}

#[tokio::main]
async fn main() {
    let excel_sheet = match parse::get_sheet("СЗМК1.xlsx", "Sheet1") {
        Ok(sheet) => Ok(sheet),
        Err(err) => {
            write_log(err.as_str());
            Err(())
        }
    }; 
    
    pretty_env_logger::init();
    log::info!("Starting bot...");

    let bot = Bot::from_env();

    Dispatcher::builder(bot, schema())
        .dependencies(dptree::deps![InMemStorage::<State>::new(), excel_sheet.unwrap()])
        .enable_ctrlc_handler()
        .build()
        .dispatch()
        .await;
}

fn schema() -> UpdateHandler<Box<dyn std::error::Error + Send + Sync + 'static>> {
    use dptree::case;

    let command_handler = teloxide::filter_command::<Command, _>()
        .branch(case![State::Start]
            .branch(case![Command::Start].endpoint(start)));

    let message_handler = Update::filter_message()
        .branch(command_handler)
        .branch(case![State::CablesRecieveNumber].endpoint(cables_receive_number))
        .branch(dptree::endpoint(invalid_state));

    let callback_query_handler = Update::filter_callback_query()
    .branch(
        case![State::InlineRecieveMode].endpoint(receive_choosed_mode),
    );

    dialogue::enter::<Update, InMemStorage<State>, State, _>()
        .branch(message_handler)
        .branch(callback_query_handler)
}

async fn start(bot: Bot, dialogue: MyDialogue) -> HandlerResult {
    let mode_buttons = ["Cables"]
        .map(|mode_button| InlineKeyboardButton::callback(mode_button, mode_button));

    bot.send_message(dialogue.chat_id(), "Select the mode:")
        .reply_markup(InlineKeyboardMarkup::new([mode_buttons]))
        .await?;

    dialogue.update(State::InlineRecieveMode).await?;
    Ok(())
}

async fn invalid_state(bot: Bot, msg: Message) -> HandlerResult {
    bot.send_message(msg.chat.id, "Unable to handle the message. Type /help to see the usage.")
        .await?;
    Ok(())
}

async fn cables_receive_number(
    bot: Bot,
    dialogue: MyDialogue,
    msg: Message,
    excel_sheet: Range<Data>
) -> HandlerResult {

    let mut messages: Vec<String> = vec![String::new()];
    let mut counter = 0;
    let cables = parse::get_all_cables_in_room(msg.text().unwrap().to_string(), &excel_sheet); 

    match cables {
        Ok(_) => {
            for cable in cables.unwrap() {
                /* 
                    cause of telegram limitation of characters on one 
                    message, need to cut cables vector to 20 lines messages
                */
                if counter >= 20 {
                    counter = 0;
                    let new_message: String = String::new();
                    messages.push(new_message);    
                }
                let len = &messages.len();
                messages[*len - 1].push_str(&format!("{} is {} \n", cable.index, cable.cable_type));
                counter += 1;
            }
            for message in messages {
                bot.send_message(dialogue.chat_id(), message).await?;
            }
            /*
                All cables found and sended, starts dialogue again 
            */
            let mode_buttons = ["Cables"]
                .map(|mode_button| InlineKeyboardButton::callback(mode_button, mode_button));

            bot.send_message(dialogue.chat_id(), "Select the mode:")
                .reply_markup(InlineKeyboardMarkup::new([mode_buttons]))
                .await?;

            dialogue.update(State::InlineRecieveMode).await?;
            return Ok(());
        }
        Err(_) => {
            /*
                If no one cable found, starts dialogue again 
             */
            write_log("Совпадений не найдено, попробуйте поменять буквы с латинских на русские, или наоборот");
            bot.send_message(msg.chat.id, "Совпадений не найдено, попробуйте поменять буквы с латинских на русские, или наоборот").await?;
            let mode_buttons = ["Cables"]
                .map(|mode_button| InlineKeyboardButton::callback(mode_button, mode_button));

            bot.send_message(dialogue.chat_id(), "Select the mode:")
                .reply_markup(InlineKeyboardMarkup::new([mode_buttons]))
                .await?;

            dialogue.update(State::InlineRecieveMode).await?;
            return Ok(());
        }
    }
}

async fn receive_choosed_mode(
    bot: Bot,
    dialogue: MyDialogue,
    q: CallbackQuery,
) -> HandlerResult {
    if let Some(mode) = &q.data {
        match mode.as_str() {
            "Cables" => {
                bot.send_message(dialogue.chat_id(), "Укажите номер помещения в 5-значном формате (прим. 07817, 10115, и т.д.)").await?;
                dialogue.update(State::CablesRecieveNumber).await?;
            }
            _ => {
                bot.send_message(dialogue.chat_id(), "Unknow query").await?;
                dialogue.exit().await?;
            }
        }
    }
    Ok(())
}