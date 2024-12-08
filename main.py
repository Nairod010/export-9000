from typing import Final
import os
from dotenv import load_dotenv
import discord
from discord import Intents, Client, Message
from discord.ext import commands
#from responses import get_response
from openpyxl import Workbook

load_dotenv()
TOKEN: Final[str] = os.getenv('DISCORD_TOKEN')

intents: Intents = Intents.default()
intents.message_content = True  #NOQA
intents.members = True #NOQA
client: Client = Client(intents=intents)
bot = commands.Bot(command_prefix='!', intents=intents)


active_survey = {"question": None, "responses": {}, "is_active": False, "message_id": None}

@bot.command()
async def start_survey(ctx,channel:discord.TextChannel = None, *, question: str):
    if not channel:
        channel = ctx.channel

    if not channel.permissions_for(ctx.guild.me).send_messages:
        await ctx.send(f"I don't have permission to send messages in {channel.mention}")
        return
    

    if not ctx.guild:
        await ctx.send("This command can only be used in a server.")
        return

    if active_survey["is_active"]:
        await ctx.send("A survey is already active. Use !finish_survey to end it.") 
        return

    active_survey.update({
        "is_active": True,
        "question": question,
        "responses": {}
    })

    try:
    
        survey_message = await channel.send(f"**Survey:** @everyone {question}\nReact with 👍 (Yes), 👎 (No), or 🤷 (IDK).")

        active_survey["message_id"] = survey_message.id

        await survey_message.add_reaction("👍")
        await survey_message.add_reaction("👎")
        await survey_message.add_reaction("🤷")

        await ctx.send("Survey started! Members can now respond.")
    except Exception as e:
        await ctx.send(f"Failed to start the survey in {channel.mention}. Error: {e}")
@start_survey.error
async def start_survey_error(ctx, error):
    if isinstance(error, commands.ChannelNotFound):
        await ctx.send("Specify the channel before asking the question. The channel mentioned was not found.")
    else:
        await ctx.send(f"An error occurred: {error}")
@bot.event
async def on_reaction_add(reaction, user):
    if user.bot or not active_survey["is_active"]:
        return

    if reaction.message.id != active_survey.get("message_id"):
        return

    user_id = user.id
    current_reactions = [
        r.emoji for r in reaction.message.reactions if r.count > 1 and str(r.emoji) in ["👍", "👎", "🤷"]
    ]

    if not current_reactions:
        active_survey["response"].pop(user_id, None)
    elif len(current_reactions) > 1:
        active_survey["responses"][user_id] = {
            "name": user.name,
            "response": "IDK",
        }
    elif "👍" in current_reactions:
        active_survey["responses"][user_id] = {
            "name": user.name,
            "response": "Yes",
        }
    elif "👎" in current_reactions:
        active_survey["responses"][user_id] = {
            "name": user.name,
            "response": "No",
        }
    elif "🤷" in current_reactions:
        active_survey["responses"][user_id] = {
            "name": user.name,
            "response": "IDK",
        }

@bot.command()
async def finish_survey(ctx):
    if not active_survey["is_active"]:
        await ctx.send("There is no active survey to finish.")
        return

    question = active_survey["question"]
    responses = active_survey["responses"]

    filename = "survey_results.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "Survey Results"

    ws.append(["Username","Response"])

    for user_id, data in active_survey["responses"].items():
        ws.append([
            data["name"],
            data["response"]
        ])
    
    wb.save(filename)

    active_survey.update({
        "is_active": False,
        "question": None,
        "responses": {},
        "message_id": None,
    })

    try:
        await ctx.author.send(f"Results for:**{question}** are attached here", file=discord.File(filename))
        await ctx.send("Survey completed!")
    except discord.Forbidden:
        await ctx.send("Survey could not be exported due to privacy settings.")
    except Exception as e:
        await ctx.send(f"An error occurred while sending the results: {e}")
    
    os.remove(filename)

@bot.command()
async def export_members(ctx):
    if not ctx.guild:
        await ctx.send("This command can only be used in a server.")
        return
    
    wb = Workbook()
    ws = wb.active
    ws.title = "Server Members"

    ws.append(["ID","Username","Discriminator","Display Name","Joined At"])

    for member in ctx.guild.members:
        joined_at = member.joined_at.strftime('%Y-%m-%d %H:%M:%S') if member.joined_at else "N/A"
        ws.append([
            member.id,
            member.name,
            member.discriminator,
            member.display_name,
            joined_at
        ])
    
    filename = "members.xlsx"
    wb.save(filename)

    await ctx.send(file=discord.File(filename))


if __name__ == '__main__':
    bot.run(token=TOKEN)
