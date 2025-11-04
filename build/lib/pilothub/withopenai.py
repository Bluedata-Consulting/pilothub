from openai import OpenAI
import openai
import time


class ContentClient:
    def __init__(self,
                 *,
                 openai_api_key: str = None,
                 open_ai_model: str = "gpt-3.5-turbo",
                 max_tokens: int = 2000,
                 temperature: float = 0.7,
                 ):
        if openai_api_key is None:
            raise ValueError("Please provide an OpenAI API key.")
        self.openai_api_key = openai_api_key
        self.openai_client = OpenAI(api_key=self.openai_api_key)
        self.open_ai_model = open_ai_model
        self.openai_max_tokens = max_tokens
        self.openai_temperature = temperature
        self.set_system_role()

    def set_system_role(self,
                        system_role: str | None = None):
        """
        Set the system role.
        :param system_role: System role to set.
        """
        if system_role is None:
            self.system_role = (
                "You are an expert in the field of "
                "information technology. You need to provide detailed notes, "
                "summary and key points for the topics provided."
            )
        else:
            self.system_role = system_role
        self.openai_message = [{"role": "system", "content": self.system_role},]

    def get_notes_from_text(self, openai_model: str | None = None,
                            openai_max_tokens: int | None = None,
                            openai_temperature: float | None = None,
                            text: str | None = None,
                            prompt: str | None = None,
                            ):
        """
        Get notes from text using OpenAI's GPT-3 model.
        :param openai_model: OpenAI model to use. Defaults to gpt-3.5-turbo.
        :param openai_max_tokens: Maximum number of tokens to generate.
                                Defaults to 2000.
        :param openai_temperature: Controls randomness. Defaults to 0.7.
        :param text: Text to generate notes from.
        :param prompt: Prompt to use for generating notes.
        :return: Notes generated from text.
        """
        if openai_model is None:
            openai_model = self.open_ai_model
        if openai_max_tokens is None:
            openai_max_tokens = self.openai_max_tokens
        if openai_temperature is None:
            openai_temperature = self.openai_temperature
        if text is None:
            raise ValueError("Please provide text to generate notes from.")
        if prompt is None:
            prompt = """
            Role: You are an expert in course creation and articulation.
            Task: You need to provide detailed notes, summary & key points.
            Contraints: Keep the notes provided precise and concise.
            Use bullet points to explain the key concepts.
            Add a section "For Students: " where provide expalanation is
                short format in layman language.(only if necessary/suitable)
            Provide notes for the following information:

            """
        prompt += "\n```" + text + "```"
        self.openai_message.append({"role": "user", "content": prompt})

        try:
            response = self.openai_client.chat.completions.create(
                model=openai_model,
                messages=self.openai_message,
                max_tokens=openai_max_tokens,
                temperature=openai_temperature)
        except openai.RateLimitError as E:
            print("Rate Limit Error ", E)
            time.sleep(65)
            response = self.openai_client.chat.completions.create(
                model=openai_model,
                messages=self.openai_message,
                max_tokens=openai_max_tokens,
                temperature=openai_temperature)
        return response.choices[0].message.content
    